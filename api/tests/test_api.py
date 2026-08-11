import asyncio
import base64
import hashlib
import importlib
import io
import json
import os
import tempfile
import time
import unittest
from pathlib import Path

import fitz
from docx import Document
from fastapi.testclient import TestClient
from openpyxl import Workbook
from pptx import Presentation


class ApiRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory(prefix="ppt-ai-api-test-")
        os.environ["STORAGE_BACKEND"] = "local"
        os.environ["DATA_DIR"] = cls.tmp.name
        os.environ.pop("ZHIPU_API_KEY", None)
        salt = b"ppt-ai-test-salt"
        digest = hashlib.pbkdf2_hmac("sha256", b"PPT-TEST-CODE", salt, 210000)
        encode = lambda value: base64.urlsafe_b64encode(value).decode("ascii").rstrip("=")
        os.environ["INVITE_CODES_JSON"] = json.dumps({
            "invite01": f"pbkdf2_sha256$210000${encode(salt)}${encode(digest)}"
        })
        os.environ["INVITE_EXPIRES_AT"] = "2099-01-01T00:00:00Z"
        os.environ["ACCESS_SESSION_SECONDS"] = str(30 * 24 * 60 * 60)
        os.environ["ACCESS_AUTH_SECRET"] = "unit-test-auth-secret"

        import main

        cls.main = importlib.reload(main)
        cls.client = TestClient(cls.main.app)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_project_file_and_persistence_flow(self):
        health = self.client.get("/health")
        self.assertEqual(health.status_code, 200)
        self.assertTrue(health.json()["storageOk"])
        self.assertTrue(health.json()["accessControlReady"])

        denied_create = self.client.post("/projects", json={"title": "Denied"})
        self.assertEqual(denied_create.status_code, 401)

        denied_login = self.client.post(
            "/auth/login",
            json={"code": "PPT-WRONG-CODE"},
        )
        self.assertEqual(denied_login.status_code, 401)

        logged_in = self.client.post(
            "/auth/login",
            json={"code": "ppt-test-code"},
        )
        self.assertEqual(logged_in.status_code, 200)
        access_token = logged_in.json()["token"]
        access_headers = {"Authorization": f"Bearer {access_token}"}
        session = self.client.get("/auth/session", headers=access_headers)
        self.assertEqual(session.status_code, 200)
        self.assertEqual(session.json()["user"]["username"], "invite01")

        created = self.client.post(
            "/projects",
            headers=access_headers,
            json={
                "title": "PPT 测试项目",
                "purpose": "验证资料顺序锁与持久化",
                "audience": "测试人员",
                "limit": "8 页",
            },
        )
        self.assertEqual(created.status_code, 200)
        body = created.json()
        self.assertEqual(body["project"]["createdBy"], "invite01")
        project_id = body["project"]["id"]
        token = body["accessToken"]
        headers = {"X-Project-Token": token}

        unauthorized = self.client.get(f"/projects/{project_id}")
        self.assertEqual(unauthorized.status_code, 401)

        uploaded = self.client.post(
            f"/projects/{project_id}/files",
            headers=headers,
            files=[
                ("files", ("第一份.txt", "第一份资料：市场规模 100。".encode("utf-8"), "text/plain")),
                ("files", ("第二份.txt", "第二份资料：客户满意度 95%。".encode("utf-8"), "text/plain")),
            ],
        )
        self.assertEqual(uploaded.status_code, 200)
        materials = uploaded.json()["project"]["materials"]
        self.assertEqual(len(materials), 2)

        first_id = materials[0]["id"]
        second_id = materials[1]["id"]
        patched = self.client.patch(
            f"/projects/{project_id}",
            headers=headers,
            json={
                "locked": True,
                "materials": [
                    {"id": first_id, "customerOrder": 2},
                    {"id": second_id, "customerOrder": 1},
                ],
            },
        )
        self.assertEqual(patched.status_code, 200)
        patched_project = patched.json()["project"]
        self.assertTrue(patched_project["locked"])
        self.assertEqual([m["id"] for m in patched_project["materials"]], [second_id, first_id])

        downloaded = self.client.get(
            f"/projects/{project_id}/files/{second_id}", headers=headers
        )
        self.assertEqual(downloaded.status_code, 200)
        self.assertIn("客户满意度", downloaded.content.decode("utf-8"))

        self.main.storage = self.main.LocalStorage(Path(self.tmp.name))
        restored = self.client.get(f"/projects/{project_id}", headers=headers)
        self.assertEqual(restored.status_code, 200)
        self.assertTrue(restored.json()["project"]["locked"])
        self.assertEqual(len(restored.json()["project"]["materials"]), 2)

        no_key = self.client.post(f"/projects/{project_id}/analyze", headers=headers)
        self.assertEqual(no_key.status_code, 503)

    def test_supported_document_parsers(self):
        pdf = fitz.open()
        page = pdf.new_page()
        page.insert_text((72, 72), "PDF parser test")
        pdf_text, previews = self.main.parse_pdf(pdf.tobytes())
        self.assertIn("PDF parser test", pdf_text)
        self.assertEqual(len(previews), 1)

        docx_buffer = io.BytesIO()
        document = Document()
        document.add_paragraph("DOCX parser test")
        document.save(docx_buffer)
        self.assertIn("DOCX parser test", self.main.parse_docx(docx_buffer.getvalue()))

        xlsx_buffer = io.BytesIO()
        workbook = Workbook()
        workbook.active["A1"] = "XLSX parser test"
        workbook.save(xlsx_buffer)
        self.assertIn("XLSX parser test", self.main.parse_xlsx(xlsx_buffer.getvalue()))

        pptx_buffer = io.BytesIO()
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[1])
        slide.shapes.title.text = "PPTX parser test"
        presentation.save(pptx_buffer)
        self.assertIn("PPTX parser test", self.main.parse_pptx(pptx_buffer.getvalue()))

    def test_access_token_never_outlives_invite(self):
        original_expiry = self.main.INVITE_EXPIRES_AT_TIMESTAMP
        try:
            self.main.INVITE_EXPIRES_AT_TIMESTAMP = int(time.time()) + 60
            token, expires_at = self.main.issue_access_token("invite01")
            encoded = token.split(".", 1)[0]
            payload = json.loads(self.main.b64url_decode(encoded).decode("utf-8"))
            self.assertEqual(payload["exp"], self.main.INVITE_EXPIRES_AT_TIMESTAMP)
            self.assertEqual(
                expires_at,
                self.main.datetime.fromtimestamp(
                    self.main.INVITE_EXPIRES_AT_TIMESTAMP,
                    tz=self.main.timezone.utc,
                ).isoformat(),
            )
        finally:
            self.main.INVITE_EXPIRES_AT_TIMESTAMP = original_expiry

    def test_ai_retries_transient_provider_overload(self):
        class FakeResponse:
            def __init__(self, status_code, payload, headers=None):
                self.status_code = status_code
                self._payload = payload
                self.headers = headers or {}
                self.text = json.dumps(payload, ensure_ascii=False)

            def json(self):
                return self._payload

        class FakeClient:
            calls = 0

            def __init__(self, timeout):
                self.timeout = timeout

            async def __aenter__(self):
                return self

            async def __aexit__(self, exc_type, exc, tb):
                return False

            async def post(self, url, headers, json):
                FakeClient.calls += 1
                if FakeClient.calls == 1:
                    return FakeResponse(429, {"error": {"code": "1305"}}, {"Retry-After": "1"})
                return FakeResponse(200, {"choices": [{"message": {"content": "ok"}}]})

        original_client = self.main.httpx.AsyncClient
        original_key = self.main.ZHIPU_API_KEY
        original_attempts = self.main.AI_MAX_ATTEMPTS
        original_sleep = self.main.asyncio.sleep
        try:
            self.main.httpx.AsyncClient = FakeClient
            self.main.ZHIPU_API_KEY = "test-key"
            self.main.AI_MAX_ATTEMPTS = 3

            async def no_sleep(_delay):
                return None

            self.main.asyncio.sleep = no_sleep
            result = asyncio.run(self.main.zhipu_chat({"model": "test"}))
            self.assertEqual(result["choices"][0]["message"]["content"], "ok")
            self.assertEqual(FakeClient.calls, 2)
        finally:
            self.main.httpx.AsyncClient = original_client
            self.main.ZHIPU_API_KEY = original_key
            self.main.AI_MAX_ATTEMPTS = original_attempts
            self.main.asyncio.sleep = original_sleep


if __name__ == "__main__":
    unittest.main()
