import argparse
import json
import os
from pathlib import Path

import httpx


def create_fixture(client: httpx.Client, state_path: Path, invite_code: str) -> None:
    logged_in = client.post(
        "/auth/login",
        json={"code": invite_code},
    )
    logged_in.raise_for_status()
    access_headers = {"Authorization": f"Bearer {logged_in.json()['token']}"}
    session = client.get("/auth/session", headers=access_headers)
    session.raise_for_status()
    assert session.json()["user"]["username"].startswith("invite")

    created = client.post(
        "/projects",
        headers=access_headers,
        json={
            "title": "部署持久化测试",
            "purpose": "验证上传、顺序锁、重新部署和 AI 编排",
            "audience": "内部产品验收",
            "limit": "6-8 页",
            "instruction": "只使用资料中的数字，不补充外部事实。",
        },
    )
    created.raise_for_status()
    payload = created.json()
    project_id = payload["project"]["id"]
    token = payload["accessToken"]
    headers = {"X-Project-Token": token}

    unauthorized = client.get(f"/projects/{project_id}")
    assert unauthorized.status_code == 401, unauthorized.text

    uploaded = client.post(
        f"/projects/{project_id}/files",
        headers=headers,
        files=[
            (
                "files",
                (
                    "01-市场资料.txt",
                    "2026 年测试市场规模为 100 万元，样本仅用于系统验收。".encode("utf-8"),
                    "text/plain",
                ),
            ),
            (
                "files",
                (
                    "02-客户反馈.txt",
                    "测试客户满意度为 95%，最关注资料顺序与事实准确性。".encode("utf-8"),
                    "text/plain",
                ),
            ),
        ],
    )
    uploaded.raise_for_status()
    materials = uploaded.json()["project"]["materials"]
    assert len(materials) == 2

    patched = client.patch(
        f"/projects/{project_id}",
        headers=headers,
        json={
            "locked": True,
            "materials": [
                {"id": materials[0]["id"], "customerOrder": 2, "weight": "重要"},
                {"id": materials[1]["id"], "customerOrder": 1, "weight": "核心"},
            ],
        },
    )
    patched.raise_for_status()
    ordered = patched.json()["project"]["materials"]
    assert ordered[0]["id"] == materials[1]["id"]
    assert patched.json()["project"]["locked"] is True

    state_path.parent.mkdir(parents=True, exist_ok=True)
    state_path.write_text(
        json.dumps(
            {
                "projectId": project_id,
                "token": token,
                "orderedFileIds": [item["id"] for item in ordered],
                "orderedNames": [item["name"] for item in ordered],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(json.dumps({"created": True, "projectId": project_id, "files": len(ordered)}, ensure_ascii=False))


def verify_fixture(client: httpx.Client, state_path: Path) -> None:
    state = json.loads(state_path.read_text(encoding="utf-8"))
    project_id = state["projectId"]
    headers = {"X-Project-Token": state["token"]}

    restored = client.get(f"/projects/{project_id}", headers=headers)
    restored.raise_for_status()
    project = restored.json()["project"]
    assert project["locked"] is True
    assert [item["id"] for item in project["materials"]] == state["orderedFileIds"]

    for material in project["materials"]:
        downloaded = client.get(f"/projects/{project_id}/files/{material['id']}", headers=headers)
        downloaded.raise_for_status()
        assert downloaded.content
        assert "filename*=UTF-8''" in downloaded.headers["content-disposition"]

    cors = client.options(
        f"/projects/{project_id}",
        headers={
            "Origin": "https://ppt-material-manager-ppt-ai-d8gx2g506c70eb54e.webapps.tcloudbase.com",
            "Access-Control-Request-Method": "GET",
            "Access-Control-Request-Headers": "x-project-token",
        },
    )
    cors.raise_for_status()
    assert cors.headers.get("access-control-allow-origin") == "https://ppt-material-manager-ppt-ai-d8gx2g506c70eb54e.webapps.tcloudbase.com"

    analyzed = client.post(f"/projects/{project_id}/analyze", headers=headers, timeout=300)
    if not analyzed.is_success:
        print(
            json.dumps(
                {
                    "analysisHttpStatus": analyzed.status_code,
                    "analysisError": analyzed.text[:2000],
                },
                ensure_ascii=False,
            )
        )
    analyzed.raise_for_status()
    analyzed_payload = analyzed.json()
    result = analyzed_payload["result"]
    assert analyzed_payload["project"]["analysisStatus"] == "done"
    assert result.get("outline")
    assert len(analyzed_payload.get("materials", [])) == 2

    print(
        json.dumps(
            {
                "verified": True,
                "projectId": project_id,
                "files": len(project["materials"]),
                "analysisStatus": "done",
                "outlineSections": len(result["outline"]),
                "model": result.get("model"),
            },
            ensure_ascii=False,
        )
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=("create", "verify"))
    parser.add_argument("--base-url", required=True)
    parser.add_argument("--state", required=True, type=Path)
    parser.add_argument("--invite-code", default=os.getenv("PPT_TEST_INVITE_CODE", ""))
    args = parser.parse_args()

    with httpx.Client(base_url=args.base_url.rstrip("/"), timeout=90, follow_redirects=True) as client:
        if args.mode == "create":
            if not args.invite_code:
                parser.error("create mode requires --invite-code or PPT_TEST_INVITE_CODE")
            create_fixture(client, args.state, args.invite_code)
        else:
            verify_fixture(client, args.state)


if __name__ == "__main__":
    main()
