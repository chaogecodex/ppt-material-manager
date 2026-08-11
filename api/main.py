import os
import io
import json
import base64
import hashlib
import hmac
import mimetypes
import re
import secrets
import time
import uuid
import asyncio
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional
from urllib.parse import quote

import httpx
import fitz  # PyMuPDF
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
from fastapi import FastAPI, UploadFile, File, HTTPException, Header, Body, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response


# -----------------------------
# Configuration
# -----------------------------
APP_ORIGINS = [x.strip() for x in os.getenv(
    "APP_ORIGINS",
    "https://ppt-material-manager-ppt-ai-d8gx2g506c70eb54e.webapps.tcloudbase.com,"
    "https://chaogecodex.github.io,http://localhost:8000,http://127.0.0.1:5500"
).split(",") if x.strip()]

ZHIPU_API_KEY = os.getenv("ZHIPU_API_KEY", "")
ZHIPU_ENDPOINT = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
TEXT_MODEL = os.getenv("TEXT_MODEL", "glm-4.7-flash")
VISION_MODEL = os.getenv("VISION_MODEL", "glm-4.6v-flash")

STORAGE_BACKEND = os.getenv("STORAGE_BACKEND", "local").lower().strip()  # local | cos
DATA_DIR = Path(os.getenv("DATA_DIR", "/tmp/ppt-material-manager-data"))
COS_SECRET_ID = os.getenv("COS_SECRET_ID", "")
COS_SECRET_KEY = os.getenv("COS_SECRET_KEY", "")
COS_REGION = os.getenv("COS_REGION", "")
COS_BUCKET = os.getenv("COS_BUCKET", "")

MAX_FILE_BYTES = int(os.getenv("MAX_FILE_BYTES", str(60 * 1024 * 1024)))
MAX_PROJECT_BYTES = int(os.getenv("MAX_PROJECT_BYTES", str(300 * 1024 * 1024)))
MAX_TEXT_CHARS_PER_FILE = int(os.getenv("MAX_TEXT_CHARS_PER_FILE", "70000"))
MAX_TOTAL_SUMMARY_CHARS = int(os.getenv("MAX_TOTAL_SUMMARY_CHARS", "140000"))
AI_MAX_ATTEMPTS = max(1, int(os.getenv("AI_MAX_ATTEMPTS", "4")))
# 分析在后台跑，状态落盘。若容器在分析途中重启，状态会永久停在 running——
# 超过这个秒数仍是 running 即判定为中断，允许重新发起。
ANALYSIS_STALE_SECONDS = int(os.getenv("ANALYSIS_STALE_SECONDS", "900"))

DEFAULT_INVITE_CODES_JSON = '{"invite01":"pbkdf2_sha256$210000$3TZnASM5ZDfKGwZRn3bEwg$LtS8Q6q4EJQmpyLY_8TVuEXJR2jIJq9TC4lWzpnv-NI","invite02":"pbkdf2_sha256$210000$nMIVlwrwMy3zIIaFR7Thyw$YpESb3zBNQoGCZR99OSjtd9NW0IS0NEJob5nbDXS5Zo","invite03":"pbkdf2_sha256$210000$s36EpIKco3LXaOmzpkyFtQ$Z8dYKHA0Q8jo5EtZ3rTuwEcLE5M5wZWQSX9z6Md7LyQ","invite04":"pbkdf2_sha256$210000$B3WdWT0RABMsrwum9sOaqw$D8N46G7EwPMRlvGqqPNXQwTioScMY_ESwpCaK2Mm334","invite05":"pbkdf2_sha256$210000$HwY9BB82qnxAwpT9HRQdBQ$3AGOZ-mBsOF9Fe55VH1Pkg3x3Ors-RUr_1bGWruur0s","invite06":"pbkdf2_sha256$210000$7UqIHeL5Vc_KixybigFv6w$AZBRYOGIrUKKqDpVGBvte5XrvLdgg62FrxZ9Wt6oFVY","invite07":"pbkdf2_sha256$210000$V8X5AWh8qgKQq-zn2DOmxg$gLBd7kuvbvtZiDGQX-zwXlFdZKVRXypfIFnMfp72VGk","invite08":"pbkdf2_sha256$210000$z6djPNt17WtuD-bm33hUtw$eDI5HWdjukVBw7jEylYK_DMmJst_wPDOu64s8QT01H0","invite09":"pbkdf2_sha256$210000$n8nas4hGIAbjzrSbcIhXew$vH3X2zG6zGAT1MCawFMHUHlXjWDOgCUr-T0aSZG9o9Q","invite10":"pbkdf2_sha256$210000$6XRC1ovJYBdUjKsjX3NUgw$OYxFjVGQYhMhkZAYXHT6vNN8anADSwFDtjEwyc7z-RQ"}'
INVITE_CODES_JSON = os.getenv("INVITE_CODES_JSON", DEFAULT_INVITE_CODES_JSON)
INVITE_EXPIRES_AT_RAW = os.getenv("INVITE_EXPIRES_AT", "2026-09-10T11:07:30Z")
try:
    invite_expiry = datetime.fromisoformat(INVITE_EXPIRES_AT_RAW.replace("Z", "+00:00"))
    if invite_expiry.tzinfo is None:
        invite_expiry = invite_expiry.replace(tzinfo=timezone.utc)
    INVITE_EXPIRES_AT_TIMESTAMP = int(invite_expiry.timestamp())
except (TypeError, ValueError) as exc:
    raise RuntimeError("INVITE_EXPIRES_AT must be an ISO 8601 timestamp") from exc
INVITE_EXPIRES_AT_ISO = datetime.fromtimestamp(
    INVITE_EXPIRES_AT_TIMESTAMP, tz=timezone.utc
).isoformat()
ACCESS_AUTH_SECRET = os.getenv("ACCESS_AUTH_SECRET", "") or (
    hashlib.sha256(f"ppt-access:{ZHIPU_API_KEY}".encode("utf-8")).hexdigest()
    if ZHIPU_API_KEY
    else ""
)
ACCESS_SESSION_SECONDS = max(300, int(os.getenv("ACCESS_SESSION_SECONDS", str(30 * 24 * 60 * 60))))
AUTH_FAILURE_LIMIT = max(3, int(os.getenv("AUTH_FAILURE_LIMIT", "8")))
AUTH_FAILURE_WINDOW_SECONDS = max(60, int(os.getenv("AUTH_FAILURE_WINDOW_SECONDS", "300")))

app = FastAPI(title="PPT Material Manager AI API", version="0.6.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=APP_ORIGINS,
    allow_credentials=False,
    allow_methods=["GET", "POST", "PATCH", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def make_project_id() -> str:
    return "p_" + uuid.uuid4().hex[:24]


def make_file_id() -> str:
    return "f_" + uuid.uuid4().hex[:20]


def hash_token(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def verify_token_hash(token: str, expected_hash: str) -> bool:
    return bool(token and expected_hash and hmac.compare_digest(hash_token(token), expected_hash))


def b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def b64url_decode(value: str) -> bytes:
    padding = "=" * (-len(value) % 4)
    return base64.urlsafe_b64decode((value + padding).encode("ascii"))


def load_invite_codes() -> Dict[str, str]:
    try:
        raw = json.loads(INVITE_CODES_JSON)
    except json.JSONDecodeError as exc:
        raise RuntimeError("INVITE_CODES_JSON is not valid JSON") from exc
    if not isinstance(raw, dict):
        raise RuntimeError("INVITE_CODES_JSON must be an object")

    invites: Dict[str, str] = {}
    for invite_id, encoded in raw.items():
        normalized = str(invite_id).strip().lower()
        if not re.fullmatch(r"[a-z0-9_-]{3,32}", normalized):
            raise RuntimeError("INVITE_CODES_JSON contains an invalid invite id")
        parts = str(encoded).split("$")
        if len(parts) != 4 or parts[0] != "pbkdf2_sha256":
            raise RuntimeError("INVITE_CODES_JSON contains an invalid code hash")
        try:
            iterations = int(parts[1])
            b64url_decode(parts[2])
            b64url_decode(parts[3])
        except (TypeError, ValueError) as exc:
            raise RuntimeError("INVITE_CODES_JSON contains an invalid code hash") from exc
        if iterations < 100000:
            raise RuntimeError("INVITE_CODES_JSON code hash iterations are too low")
        invites[normalized] = str(encoded)
    return invites


INVITE_CODES = load_invite_codes()
auth_failures: Dict[str, List[float]] = {}
auth_lock = asyncio.Lock()


def verify_invite_secret(code: str, encoded: str) -> bool:
    try:
        algorithm, iterations_text, salt_text, digest_text = encoded.split("$")
        if algorithm != "pbkdf2_sha256":
            return False
        salt = b64url_decode(salt_text)
        expected = b64url_decode(digest_text)
        actual = hashlib.pbkdf2_hmac(
            "sha256",
            code.encode("utf-8"),
            salt,
            int(iterations_text),
        )
        return hmac.compare_digest(actual, expected)
    except (ValueError, TypeError):
        return False


def resolve_invite_code(code: str) -> Optional[str]:
    """Check every stored hash so valid and invalid codes have similar cost."""
    matched_id: Optional[str] = None
    for invite_id, encoded in INVITE_CODES.items():
        if verify_invite_secret(code, encoded):
            matched_id = invite_id
    return matched_id


def issue_access_token(invite_id: str) -> Tuple[str, str]:
    expires_at = min(
        int(time.time()) + ACCESS_SESSION_SECONDS,
        INVITE_EXPIRES_AT_TIMESTAMP,
    )
    payload = {
        "sub": invite_id,
        "exp": expires_at,
        "nonce": secrets.token_hex(8),
    }
    encoded = b64url_encode(json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8"))
    signature = b64url_encode(hmac.new(
        ACCESS_AUTH_SECRET.encode("utf-8"),
        encoded.encode("ascii"),
        hashlib.sha256,
    ).digest())
    expires_iso = datetime.fromtimestamp(expires_at, tz=timezone.utc).isoformat()
    return f"{encoded}.{signature}", expires_iso


def require_access_user(authorization: Optional[str]) -> str:
    if not ACCESS_AUTH_SECRET or not INVITE_CODES:
        raise HTTPException(status_code=503, detail="登录服务尚未配置")
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=401, detail="请先登录")
    token = authorization[7:].strip()
    try:
        encoded, signature = token.split(".", 1)
        expected = b64url_encode(hmac.new(
            ACCESS_AUTH_SECRET.encode("utf-8"),
            encoded.encode("ascii"),
            hashlib.sha256,
        ).digest())
        if not hmac.compare_digest(signature, expected):
            raise ValueError("signature mismatch")
        payload = json.loads(b64url_decode(encoded).decode("utf-8"))
        invite_id = str(payload.get("sub") or "").lower()
        now = int(time.time())
        if (
            invite_id not in INVITE_CODES
            or INVITE_EXPIRES_AT_TIMESTAMP <= now
            or int(payload.get("exp") or 0) <= now
        ):
            raise ValueError("expired or unknown invite")
        return invite_id
    except (ValueError, TypeError, json.JSONDecodeError):
        raise HTTPException(status_code=401, detail="登录已失效，请重新登录")


def request_client_key(request: Request) -> str:
    forwarded = (request.headers.get("x-forwarded-for") or "").split(",", 1)[0].strip()
    return forwarded or (request.client.host if request.client else "unknown")


def safe_filename(name: str) -> str:
    name = (name or "unnamed").strip().replace("\\", "_").replace("/", "_")
    name = re.sub(r"[\x00-\x1f\x7f]", "_", name)
    return name[:180] or "unnamed"


def human_size(size: int) -> str:
    if size <= 0:
        return "—"
    units = ["B", "KB", "MB", "GB"]
    value = float(size)
    idx = 0
    while value >= 1024 and idx < len(units) - 1:
        value /= 1024
        idx += 1
    return f"{value:.1f} {units[idx]}" if idx else f"{int(value)} B"


# -----------------------------
# Persistent storage abstraction
# -----------------------------
class Storage:
    def put(self, key: str, data: bytes, content_type: str = "application/octet-stream") -> None:
        raise NotImplementedError

    def get(self, key: str) -> bytes:
        raise NotImplementedError

    def delete(self, key: str) -> None:
        raise NotImplementedError

    def exists(self, key: str) -> bool:
        raise NotImplementedError


class LocalStorage(Storage):
    def __init__(self, root: Path):
        self.root = root
        self.root.mkdir(parents=True, exist_ok=True)

    def _path(self, key: str) -> Path:
        p = (self.root / key).resolve()
        root = self.root.resolve()
        if root != p and root not in p.parents:
            raise ValueError("invalid storage key")
        return p

    def put(self, key: str, data: bytes, content_type: str = "application/octet-stream") -> None:
        p = self._path(key)
        p.parent.mkdir(parents=True, exist_ok=True)
        tmp = p.with_suffix(p.suffix + ".tmp")
        tmp.write_bytes(data)
        tmp.replace(p)

    def get(self, key: str) -> bytes:
        p = self._path(key)
        if not p.exists():
            raise FileNotFoundError(key)
        return p.read_bytes()

    def delete(self, key: str) -> None:
        p = self._path(key)
        try:
            p.unlink()
        except FileNotFoundError:
            pass

    def exists(self, key: str) -> bool:
        return self._path(key).exists()


class CosStorage(Storage):
    def __init__(self):
        if not all([COS_SECRET_ID, COS_SECRET_KEY, COS_REGION, COS_BUCKET]):
            raise RuntimeError("COS storage selected but COS_SECRET_ID/COS_SECRET_KEY/COS_REGION/COS_BUCKET are incomplete")
        from qcloud_cos import CosConfig, CosS3Client
        cfg = CosConfig(
            Region=COS_REGION,
            SecretId=COS_SECRET_ID,
            SecretKey=COS_SECRET_KEY,
            Scheme="https",
        )
        self.client = CosS3Client(cfg)

    def put(self, key: str, data: bytes, content_type: str = "application/octet-stream") -> None:
        self.client.put_object(Bucket=COS_BUCKET, Body=data, Key=key, ContentType=content_type)

    def get(self, key: str) -> bytes:
        try:
            obj = self.client.get_object(Bucket=COS_BUCKET, Key=key)
            return obj["Body"].get_raw_stream().read()
        except Exception as exc:
            raise FileNotFoundError(key) from exc

    def delete(self, key: str) -> None:
        try:
            self.client.delete_object(Bucket=COS_BUCKET, Key=key)
        except Exception:
            pass

    def exists(self, key: str) -> bool:
        try:
            self.client.head_object(Bucket=COS_BUCKET, Key=key)
            return True
        except Exception:
            return False


def build_storage() -> Storage:
    if STORAGE_BACKEND == "cos":
        return CosStorage()
    return LocalStorage(DATA_DIR)


storage = build_storage()
project_locks: Dict[str, asyncio.Lock] = {}
# 正在跑的分析任务。服务是单 worker（见 Dockerfile 的 --workers 1），
# 所以这张表就是"此刻是否真的在分析"的权威依据；落盘的 analysisStatus
# 是给客户端轮询用的，进程重启后会残留 running，不能用来做重复提交判断。
analysis_tasks: Dict[str, "asyncio.Task[None]"] = {}


def get_lock(project_id: str) -> asyncio.Lock:
    if project_id not in project_locks:
        project_locks[project_id] = asyncio.Lock()
    return project_locks[project_id]


def project_key(project_id: str) -> str:
    return f"projects/{project_id}/project.json"


def raw_file_key(project_id: str, file_id: str, filename: str) -> str:
    return f"projects/{project_id}/files/{file_id}/{safe_filename(filename)}"


def save_project(project: Dict[str, Any]) -> None:
    project["updatedAt"] = now_iso()
    payload = json.dumps(project, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    storage.put(project_key(project["id"]), payload, "application/json; charset=utf-8")


def load_project(project_id: str) -> Dict[str, Any]:
    try:
        raw = storage.get(project_key(project_id))
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="项目不存在或已删除")
    try:
        return json.loads(raw.decode("utf-8"))
    except Exception as exc:
        raise HTTPException(status_code=500, detail="项目元数据损坏") from exc


def public_project(project: Dict[str, Any]) -> Dict[str, Any]:
    result = json.loads(json.dumps(project, ensure_ascii=False))
    result.pop("accessTokenHash", None)
    for item in result.get("materials", []):
        item.pop("storageKey", None)
    return result


def require_project(project_id: str, token: Optional[str]) -> Dict[str, Any]:
    project = load_project(project_id)
    if not verify_token_hash(token or "", project.get("accessTokenHash", "")):
        raise HTTPException(status_code=401, detail="项目访问凭证无效")
    return project


def analysis_is_stale(project: Dict[str, Any]) -> bool:
    """running 状态是否已经不可信（进程在分析途中没了）。

    只有 running 才谈得上过期；done/failed 是终态。
    时间戳缺失或解析不了，一律当过期处理——宁可允许用户重试，
    也不要把项目永久锁在一个不会自己结束的状态里。
    """
    if project.get("analysisStatus") != "running":
        return False
    started = project.get("analysisStartedAt")
    if not started:
        return True
    try:
        started_at = datetime.fromisoformat(str(started).replace("Z", "+00:00"))
    except ValueError:
        return True
    if started_at.tzinfo is None:
        started_at = started_at.replace(tzinfo=timezone.utc)
    return (datetime.now(timezone.utc) - started_at).total_seconds() > ANALYSIS_STALE_SECONDS


# -----------------------------
# File parsing
# -----------------------------
def trim_text(text: str, limit: int = MAX_TEXT_CHARS_PER_FILE) -> str:
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    head = int(limit * 0.72)
    tail = limit - head
    return text[:head] + "\n\n[中间内容因长度限制已压缩]\n\n" + text[-tail:]


def parse_pdf(data: bytes) -> Tuple[str, List[bytes]]:
    doc = fitz.open(stream=data, filetype="pdf")
    chunks = []
    for i, page in enumerate(doc):
        txt = page.get_text("text") or ""
        chunks.append(f"\n--- PDF 第 {i+1} 页 ---\n{txt}")
    text = "".join(chunks).strip()

    preview_images: List[bytes] = []
    sparse = len(text) < max(500, len(doc) * 80)
    if sparse:
        # First pages are enough for MVP fallback; avoids exploding vision cost on long scans.
        picks = list(range(min(8, len(doc))))
        for idx in picks:
            pix = doc[idx].get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
            preview_images.append(pix.tobytes("png"))
    return trim_text(text), preview_images


def parse_docx(data: bytes) -> str:
    doc = Document(io.BytesIO(data))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table_idx, table in enumerate(doc.tables, 1):
        rows = []
        for row in table.rows[:120]:
            rows.append(" | ".join(cell.text.strip() for cell in row.cells))
        if rows:
            parts.append(f"\n--- 表格 {table_idx} ---\n" + "\n".join(rows))
    return trim_text("\n".join(parts))


def parse_xlsx(data: bytes) -> str:
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    parts = []
    for ws in wb.worksheets[:20]:
        parts.append(f"\n--- 工作表：{ws.title} ---")
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            if r_idx > 300:
                parts.append("[该工作表其余行已压缩]")
                break
            values = ["" if v is None else str(v) for v in row[:40]]
            if any(v.strip() for v in values):
                parts.append("\t".join(values))
    return trim_text("\n".join(parts))


def parse_pptx(data: bytes) -> str:
    prs = Presentation(io.BytesIO(data))
    parts = []
    for idx, slide in enumerate(prs.slides, 1):
        texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text and shape.text.strip():
                texts.append(shape.text.strip())
        if texts:
            parts.append(f"\n--- PPT 第 {idx} 页 ---\n" + "\n".join(texts))
    return trim_text("\n".join(parts))


def parse_textlike(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
        try:
            return trim_text(data.decode(enc))
        except UnicodeDecodeError:
            continue
    return trim_text(data.decode("utf-8", errors="replace"))


def detect_kind(filename: str, content_type: str) -> str:
    name = (filename or "").lower()
    content_type = (content_type or "").lower()
    if content_type.startswith("image/") or name.endswith((".png", ".jpg", ".jpeg", ".webp")):
        return "image"
    if content_type.startswith("video/") or name.endswith((".mp4", ".mov", ".m4v")):
        return "video"
    if content_type.startswith("audio/") or name.endswith((".mp3", ".wav", ".m4a")):
        return "audio"
    if name.endswith(".pdf"):
        return "pdf"
    if name.endswith(".docx"):
        return "docx"
    if name.endswith((".xlsx", ".xlsm")):
        return "xlsx"
    if name.endswith(".pptx"):
        return "pptx"
    if name.endswith((".txt", ".md", ".csv")):
        return "text"
    return "unknown"


# -----------------------------
# Zhipu AI
# -----------------------------
def require_ai_key() -> None:
    if not ZHIPU_API_KEY:
        raise HTTPException(status_code=503, detail="AI后端已运行，但还没有配置 ZHIPU_API_KEY")


async def zhipu_chat(payload: Dict[str, Any], timeout: float = 150.0) -> Dict[str, Any]:
    require_ai_key()
    headers = {"Authorization": f"Bearer {ZHIPU_API_KEY}", "Content-Type": "application/json"}
    async with httpx.AsyncClient(timeout=timeout) as client:
        for attempt in range(1, AI_MAX_ATTEMPTS + 1):
            response = await client.post(ZHIPU_ENDPOINT, headers=headers, json=payload)
            if response.status_code < 400:
                return response.json()

            retryable = response.status_code in {429, 500, 502, 503, 504}
            if not retryable or attempt == AI_MAX_ATTEMPTS:
                detail = response.text[:1200]
                raise HTTPException(
                    status_code=502,
                    detail=f"智谱 API 错误 {response.status_code}（尝试 {attempt}/{AI_MAX_ATTEMPTS}）: {detail}",
                )

            retry_after = response.headers.get("Retry-After", "").strip()
            try:
                delay = max(1.0, min(float(retry_after), 30.0)) if retry_after else min(2 ** attempt, 16)
            except ValueError:
                delay = min(2 ** attempt, 16)
            await asyncio.sleep(delay)

    raise HTTPException(status_code=502, detail="智谱 API 调用未返回结果")


async def summarize_image(material_id: str, name: str, data: bytes, mime: str) -> Dict[str, Any]:
    b64 = base64.b64encode(data).decode("ascii")
    data_url = f"data:{mime or 'image/png'};base64,{b64}"
    payload = {
        "model": VISION_MODEL,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": data_url}},
                {"type": "text", "text": (
                    "你正在为PPT资料管家分析客户上传素材。只描述图中可见事实，不猜测。"
                    "请输出简洁中文：图像内容、可用于PPT的价值、适合的页面用途、不确定点。"
                )},
            ],
        }],
        "thinking": {"type": "enabled"},
        "max_tokens": 2500,
    }
    result = await zhipu_chat(payload)
    content = result["choices"][0]["message"].get("content", "")
    return {
        "materialId": material_id,
        "name": name,
        "kind": "image",
        "summary": content,
        "keyFacts": [],
        "suggestedUse": [],
        "risks": [],
        "confidence": 0.8,
        "model": VISION_MODEL,
    }


async def summarize_scanned_pdf(material_id: str, name: str, images: List[bytes]) -> str:
    summaries = []
    for idx, img in enumerate(images, 1):
        item = await summarize_image(material_id, f"{name} / 页面{idx}", img, "image/png")
        summaries.append(item["summary"])
    return "\n\n".join(summaries)


async def summarize_text_material(material_id: str, name: str, kind: str, text: str, weight: str, customer_order: int) -> Dict[str, Any]:
    system = (
        "你是PPT资料编排分析器。准确提炼素材，不自由改写客户事实。"
        "不得编造输入中不存在的信息。客户顺序是硬约束，不得自行改变。"
        "返回JSON对象，字段：summary, keyFacts, suggestedUse, risks, confidence。"
    )
    user = f"""素材名称：{name}
素材类型：{kind}
客户顺序：{customer_order}
客户权重：{weight}

素材内容：
{text}
"""
    payload = {
        "model": TEXT_MODEL,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "response_format": {"type": "json_object"},
        "thinking": {"type": "enabled"},
        "max_tokens": 5000,
    }
    result = await zhipu_chat(payload)
    raw = result["choices"][0]["message"].get("content", "{}")
    try:
        parsed = json.loads(raw)
    except Exception:
        parsed = {
            "summary": raw,
            "keyFacts": [],
            "suggestedUse": [],
            "risks": ["模型未返回标准JSON"],
            "confidence": 0.5,
        }
    parsed.update({"materialId": material_id, "name": name, "kind": kind, "model": TEXT_MODEL})
    return parsed


def material_prompt_block(item: Dict[str, Any], meta: Dict[str, Any]) -> str:
    return (
        f"\n### 素材 {meta['customerOrder']}｜{meta['name']}\n"
        f"素材ID：{meta['id']}\n"
        f"权重：{meta['weight']}｜类型：{meta['type']}\n"
        f"摘要：{item.get('summary') or ''}\n"
        f"关键事实：{json.dumps(item.get('keyFacts') or [], ensure_ascii=False)}\n"
        f"建议用途：{json.dumps(item.get('suggestedUse') or [], ensure_ascii=False)}\n"
        f"风险/不确定：{json.dumps(item.get('risks') or [], ensure_ascii=False)}\n"
    )


async def synthesize_outline(project: Dict[str, Any], analyses: List[Dict[str, Any]]) -> Dict[str, Any]:
    meta_by_id = {m["id"]: m for m in project.get("materials", [])}
    blocks = []
    for item in analyses:
        meta = meta_by_id.get(item.get("materialId"), {
            "id": item.get("materialId", "unknown"),
            "name": item.get("name", "unnamed"),
            "type": item.get("kind", "FILE"),
            "weight": "重要",
            "customerOrder": 999,
        })
        blocks.append(material_prompt_block(item, meta))
    material_text = trim_text("\n".join(blocks), MAX_TOTAL_SUMMARY_CHARS)

    brief = project.get("brief", {})
    system = """你是“PPT资料管家”的编排Agent。必须遵守：
1. 客户指定的素材顺序是最终编排锁。可以在每份素材内部拆页、合并重复信息，但不得跨素材调换先后。
2. 权重：核心=必须突出；重要=应采用；参考=按价值择取；忽略=不得使用。
3. 只使用素材中明确存在的事实，不补写外部知识，不猜测。
4. 输出给客户审查的是宏观编排，不是逐页文案。
5. 冲突、缺口、扫描失败、媒体暂未解析放入 conflicts，不自行填补。
6. 返回严格JSON，不要附加解释。

JSON结构：
{
  "pageEstimate": "15-18页",
  "timeEstimate": "20分钟",
  "usedCount": 3,
  "model": "glm-4.7-flash",
  "outline": [
    {"title":"章节标题","summary":"本章讲什么","pages":3,"sourceNames":["文件名"],"sourceIds":["素材ID"],"tags":["核心","数据图表"]}
  ],
  "conflicts": ["问题"],
  "orderSuggestions": ["只给建议，不自动调整客户顺序"]
}
"""
    user = f"""项目目的：{brief.get('purpose','')}
主要受众：{brief.get('audience','')}
页数/时长：{brief.get('limit','')}
客户补充要求：{brief.get('instruction','')}

以下素材已按客户顺序排列：
{material_text}
"""
    payload = {
        "model": TEXT_MODEL,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "response_format": {"type": "json_object"},
        "thinking": {"type": "enabled"},
        "max_tokens": 9000,
    }
    result = await zhipu_chat(payload)
    raw = result["choices"][0]["message"].get("content", "{}")
    try:
        parsed = json.loads(raw)
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"AI返回的编排不是有效JSON: {exc}")
    parsed.setdefault("model", TEXT_MODEL)
    return parsed


async def analyze_material(material: Dict[str, Any]) -> Dict[str, Any]:
    data = storage.get(material["storageKey"])
    material_id = material["id"]
    name = material["name"]
    kind = material["kind"]
    weight = material.get("weight", "重要")
    order = int(material.get("customerOrder", 999))

    if weight == "忽略":
        return {
            "materialId": material_id,
            "name": name,
            "kind": kind,
            "summary": "客户标记为忽略，未送AI解析",
            "keyFacts": [],
            "suggestedUse": [],
            "risks": [],
            "confidence": 1,
        }

    if kind == "image":
        return await summarize_image(material_id, name, data, material.get("mime") or "image/png")

    if kind in {"video", "audio"}:
        return {
            "materialId": material_id,
            "name": name,
            "kind": kind,
            "summary": "当前MVP尚未对视频/音频做真实内容解析。",
            "keyFacts": [],
            "suggestedUse": ["保留媒体位置，待下一版解析"],
            "risks": ["媒体内容未解析"],
            "confidence": 0,
        }

    preview_images: List[bytes] = []
    if kind == "pdf":
        text, preview_images = parse_pdf(data)
    elif kind == "docx":
        text = parse_docx(data)
    elif kind == "xlsx":
        text = parse_xlsx(data)
    elif kind == "pptx":
        text = parse_pptx(data)
    elif kind == "text":
        text = parse_textlike(data)
    else:
        return {
            "materialId": material_id,
            "name": name,
            "kind": kind,
            "summary": "暂不支持该文件格式的内容解析。",
            "keyFacts": [],
            "suggestedUse": [],
            "risks": ["不支持的文件格式"],
            "confidence": 0,
        }

    if preview_images:
        vision_summary = await summarize_scanned_pdf(material_id, name, preview_images)
        if text.strip():
            text = text + "\n\n--- 扫描页面视觉补充 ---\n" + vision_summary
        else:
            text = vision_summary
        return await summarize_text_material(material_id, name, "scanned-pdf", text, weight, order)

    if not text.strip():
        return {
            "materialId": material_id,
            "name": name,
            "kind": kind,
            "summary": "文件可读取，但未提取到可分析文本。",
            "keyFacts": [],
            "suggestedUse": [],
            "risks": ["未提取到文本"],
            "confidence": 0,
        }

    return await summarize_text_material(material_id, name, kind, text, weight, order)


# -----------------------------
# API
# -----------------------------
@app.get("/")
async def root():
    return {"service": "PPT Material Manager AI API", "version": "0.6.0", "ok": True}


@app.get("/health")
async def health():
    storage_ok = True
    storage_error = None
    try:
        probe = f"_health/{uuid.uuid4().hex}.txt"
        storage.put(probe, b"ok", "text/plain")
        storage_ok = storage.get(probe) == b"ok"
        storage.delete(probe)
    except Exception as exc:
        storage_ok = False
        storage_error = str(exc)[:300]
    return {
        "ok": storage_ok,
        "aiConfigured": bool(ZHIPU_API_KEY),
        "storageBackend": STORAGE_BACKEND,
        "storageOk": storage_ok,
        "storageError": storage_error,
        "textModel": TEXT_MODEL,
        "visionModel": VISION_MODEL,
        "accessControlReady": bool(
            ACCESS_AUTH_SECRET
            and INVITE_CODES
            and INVITE_EXPIRES_AT_TIMESTAMP > int(time.time())
        ),
        "inviteCodes": len(INVITE_CODES),
        "inviteExpiresAt": INVITE_EXPIRES_AT_ISO,
    }


@app.post("/auth/login")
async def login(request: Request, payload: Dict[str, Any] = Body(...)):
    if not ACCESS_AUTH_SECRET or not INVITE_CODES:
        raise HTTPException(status_code=503, detail="登录服务尚未配置")

    code = str(payload.get("code") or "").strip().upper()
    client_key = request_client_key(request)
    now = time.time()
    if INVITE_EXPIRES_AT_TIMESTAMP <= int(now):
        raise HTTPException(status_code=401, detail="邀请码已过期")

    async with auth_lock:
        recent = [
            attempted_at
            for attempted_at in auth_failures.get(client_key, [])
            if now - attempted_at < AUTH_FAILURE_WINDOW_SECONDS
        ]
        auth_failures[client_key] = recent
        if len(recent) >= AUTH_FAILURE_LIMIT:
            raise HTTPException(status_code=429, detail="尝试次数过多，请稍后再试")

    invite_id = resolve_invite_code(code)

    if not invite_id:
        async with auth_lock:
            auth_failures.setdefault(client_key, []).append(now)
        raise HTTPException(status_code=401, detail="邀请码错误或已过期")

    async with auth_lock:
        auth_failures.pop(client_key, None)

    token, expires_at = issue_access_token(invite_id)
    return {
        "success": True,
        "token": token,
        "expiresAt": expires_at,
        "user": {"username": invite_id},
    }


@app.get("/auth/session")
async def auth_session(authorization: Optional[str] = Header(default=None)):
    username = require_access_user(authorization)
    return {"success": True, "user": {"username": username}}


@app.post("/projects")
async def create_project(
    payload: Dict[str, Any] = Body(default={}),
    authorization: Optional[str] = Header(default=None),
):
    access_user = require_access_user(authorization)
    project_id = make_project_id()
    access_token = secrets.token_urlsafe(32)
    created = now_iso()
    project = {
        "id": project_id,
        "accessTokenHash": hash_token(access_token),
        "title": (payload.get("title") or "未命名PPT项目")[:100],
        "createdAt": created,
        "updatedAt": created,
        "createdBy": access_user,
        "brief": {
            "purpose": payload.get("purpose", ""),
            "audience": payload.get("audience", ""),
            "limit": payload.get("limit", ""),
            "instruction": payload.get("instruction", ""),
        },
        "materials": [],
        "analysisStatus": "idle",
        "aiResult": None,
        "aiMaterials": [],
        "locked": False,
    }
    save_project(project)
    return {"success": True, "project": public_project(project), "accessToken": access_token}


@app.get("/projects/{project_id}")
async def get_project(project_id: str, x_project_token: Optional[str] = Header(default=None)):
    project = require_project(project_id, x_project_token)
    return {"success": True, "project": public_project(project)}


@app.patch("/projects/{project_id}")
async def update_project(
    project_id: str,
    payload: Dict[str, Any] = Body(...),
    x_project_token: Optional[str] = Header(default=None),
):
    async with get_lock(project_id):
        project = require_project(project_id, x_project_token)

        if "title" in payload:
            project["title"] = str(payload.get("title") or "未命名PPT项目")[:100]

        brief_in = payload.get("brief")
        if isinstance(brief_in, dict):
            brief = project.setdefault("brief", {})
            for key in ("purpose", "audience", "limit", "instruction"):
                if key in brief_in:
                    brief[key] = str(brief_in.get(key) or "")[:4000]

        materials_in = payload.get("materials")
        if isinstance(materials_in, list):
            existing = {m["id"]: m for m in project.get("materials", [])}
            for patch in materials_in:
                item = existing.get(str(patch.get("id", "")))
                if not item:
                    continue
                if patch.get("weight") in {"核心", "重要", "参考", "忽略"}:
                    item["weight"] = patch["weight"]
                if "customerOrder" in patch:
                    try:
                        item["customerOrder"] = max(1, int(patch["customerOrder"]))
                    except Exception:
                        pass
            project["materials"] = sorted(existing.values(), key=lambda m: (m.get("customerOrder", 999), m.get("uploadedAt", "")))

        if "locked" in payload:
            project["locked"] = bool(payload["locked"])

        # Any edit that affects content invalidates the old outline.
        if any(k in payload for k in ("brief", "materials")):
            project["analysisStatus"] = "stale" if project.get("aiResult") else "idle"

        save_project(project)
        return {"success": True, "project": public_project(project)}


@app.post("/projects/{project_id}/files")
async def upload_project_files(
    project_id: str,
    files: List[UploadFile] = File(...),
    x_project_token: Optional[str] = Header(default=None),
):
    async with get_lock(project_id):
        project = require_project(project_id, x_project_token)
        current_bytes = sum(int(m.get("sizeBytes", 0)) for m in project.get("materials", []))
        uploaded = []

        for upload in files:
            data = await upload.read()
            if len(data) > MAX_FILE_BYTES:
                raise HTTPException(status_code=413, detail=f"{upload.filename} 超过单文件大小限制")
            if current_bytes + len(data) > MAX_PROJECT_BYTES:
                raise HTTPException(status_code=413, detail="当前项目文件总量超过限制")

            file_id = make_file_id()
            name = safe_filename(upload.filename or "unnamed")
            mime = upload.content_type or mimetypes.guess_type(name)[0] or "application/octet-stream"
            kind = detect_kind(name, mime)
            key = raw_file_key(project_id, file_id, name)
            storage.put(key, data, mime)

            item = {
                "id": file_id,
                "name": name,
                "type": kind.upper() if kind not in {"image", "video", "audio"} else {"image": "IMG", "video": "VIDEO", "audio": "AUDIO"}[kind],
                "kind": kind,
                "mime": mime,
                "sizeBytes": len(data),
                "size": human_size(len(data)),
                "weight": "参考" if kind in {"image", "video", "audio"} else "重要",
                "customerOrder": len(project.get("materials", [])) + 1,
                "storageKey": key,
                "saveStatus": "saved",
                "uploadedAt": now_iso(),
            }
            project.setdefault("materials", []).append(item)
            current_bytes += len(data)
            uploaded.append(item)

        project["aiResult"] = None
        project["aiMaterials"] = []
        project["analysisStatus"] = "idle"
        save_project(project)
        safe_uploaded = [dict(x, storageKey=None) for x in uploaded]
        for x in safe_uploaded:
            x.pop("storageKey", None)
        return {"success": True, "uploaded": safe_uploaded, "project": public_project(project)}


@app.delete("/projects/{project_id}/files/{file_id}")
async def delete_project_file(
    project_id: str,
    file_id: str,
    x_project_token: Optional[str] = Header(default=None),
):
    async with get_lock(project_id):
        project = require_project(project_id, x_project_token)
        target = next((m for m in project.get("materials", []) if m.get("id") == file_id), None)
        if not target:
            raise HTTPException(status_code=404, detail="文件不存在")
        storage.delete(target.get("storageKey", ""))
        project["materials"] = [m for m in project.get("materials", []) if m.get("id") != file_id]
        for idx, item in enumerate(project["materials"], 1):
            item["customerOrder"] = idx
        project["aiResult"] = None
        project["aiMaterials"] = []
        project["analysisStatus"] = "idle"
        save_project(project)
        return {"success": True, "project": public_project(project)}


@app.get("/projects/{project_id}/files/{file_id}")
async def get_project_file(
    project_id: str,
    file_id: str,
    x_project_token: Optional[str] = Header(default=None),
):
    project = require_project(project_id, x_project_token)
    target = next((m for m in project.get("materials", []) if m.get("id") == file_id), None)
    if not target:
        raise HTTPException(status_code=404, detail="文件不存在")
    data = storage.get(target["storageKey"])
    encoded_filename = quote(safe_filename(target["name"]), safe="")
    return Response(content=data, media_type=target.get("mime") or "application/octet-stream", headers={
        "Content-Disposition": f"inline; filename=download; filename*=UTF-8''{encoded_filename}",
        "Cache-Control": "private, max-age=60",
    })


async def run_analysis(project_id: str, materials: List[Dict[str, Any]]) -> None:
    """后台执行分析，结果一律落盘。

    进程内不保留权威状态：客户端读的是存储里的 analysisStatus，
    所以即使这个任务所在的进程没了，前端也不会永远等一个内存里的 future。

    这里全程持项目锁，挡住分析期间的 PATCH / 上传 / 删除——那些操作会重置
    aiResult，与分析结果交错写入就会产生「大纲对不上素材」的脏状态。
    代价是发起分析的接口不能再争这把锁（见 analyze_project 的说明）。
    """
    async with get_lock(project_id):
        try:
            project = load_project(project_id)
        except HTTPException:
            return  # 项目在分析期间被删了，直接收工

        try:
            analyses: List[Dict[str, Any]] = []
            for item in materials:
                analyses.append(await analyze_material(item))
            result = await synthesize_outline(project, analyses)
            project["aiMaterials"] = analyses
            project["aiResult"] = result
            project["analysisStatus"] = "done"
            project["analysisCompletedAt"] = now_iso()
            project.pop("analysisError", None)
            save_project(project)
        except Exception as exc:
            # 失败原因要留给用户看，所以写进项目而不是只抛出去——
            # 此时早已没有 HTTP 响应可以承载它了。
            detail = exc.detail if isinstance(exc, HTTPException) else f"分析失败：{str(exc)[:500]}"
            project["analysisStatus"] = "failed"
            project["analysisCompletedAt"] = now_iso()
            project["analysisError"] = str(detail)[:500]
            save_project(project)


@app.post("/projects/{project_id}/analyze", status_code=202)
async def analyze_project(
    project_id: str,
    x_project_token: Optional[str] = Header(default=None),
):
    """发起分析后立即返回，不再占着连接等 AI。

    分析耗时可达数分钟，同步返回会撞上网关超时（线上实际出现过 504）。
    客户端改为轮询 GET /projects/{id}/analyze。
    """
    require_ai_key()
    # 刻意不持项目锁：run_analysis 全程持锁，这里一争锁就会等到那次分析结束，
    # 等锁等到手时状态已经是 done，"是否在跑"的判断就永远失效了。
    # 下面从校验到 create_task 之间没有 await，单 worker 事件循环里不会被插队。
    project = require_project(project_id, x_project_token)
    materials = sorted(project.get("materials", []), key=lambda m: m.get("customerOrder", 999))
    if not materials:
        raise HTTPException(status_code=400, detail="请先上传素材")

    def running_payload(proj: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "success": True,
            "analysisStatus": "running",
            "startedAt": proj.get("analysisStartedAt"),
            "project": public_project(proj),
        }

    # 重复提交（连点、刷新后重发）不该把同一批素材再打给 AI 一遍。
    # 判断依据是进程内任务表而不是落盘状态：进程重启后落盘会残留 running，
    # 那时并没有任务在跑，用户应当可以立刻重试，而不是干等到状态过期。
    existing = analysis_tasks.get(project_id)
    if existing is not None and not existing.done():
        return running_payload(project)

    project["analysisStatus"] = "running"
    project["analysisStartedAt"] = now_iso()
    project.pop("analysisCompletedAt", None)
    project.pop("analysisError", None)
    save_project(project)

    task = asyncio.create_task(run_analysis(project_id, materials))
    analysis_tasks[project_id] = task
    task.add_done_callback(lambda _: analysis_tasks.pop(project_id, None))
    return running_payload(project)


@app.get("/projects/{project_id}/analyze")
async def analyze_status(
    project_id: str,
    x_project_token: Optional[str] = Header(default=None),
):
    """轮询端点：只回状态与结果，不回整个项目。

    刻意不加项目锁——分析任务全程持锁，这里一旦争锁，轮询就会被自己等的
    那个分析卡住。这里只读，不写。
    """
    project = require_project(project_id, x_project_token)
    status = project.get("analysisStatus", "idle")
    error = project.get("analysisError")
    if analysis_is_stale(project):
        status = "failed"
        error = "分析已中断，请重试"

    payload: Dict[str, Any] = {
        "success": True,
        "analysisStatus": status,
        "startedAt": project.get("analysisStartedAt"),
        "completedAt": project.get("analysisCompletedAt"),
    }
    if status == "done":
        # 带上 project，形状与旧的同步响应一致，客户端拿到结果后不必再多发一次请求
        payload["result"] = project.get("aiResult")
        payload["materials"] = project.get("aiMaterials", [])
        payload["project"] = public_project(project)
    elif status == "failed":
        payload["error"] = error or "分析失败，请重试"
    return payload
