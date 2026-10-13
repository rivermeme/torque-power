"""Streamlit web app: Screenshot -> motor specs -> diagram -> formula-based Excel.
Run: streamlit run "torque power app.py"
"""
from __future__ import annotations

import base64
import io
import hashlib
import json
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path

from streamlit_paste_button import paste_image_button
import requests
import streamlit as st

from motor_core import (
    calc_at, calc_switched_at, generate_excel, group_rows, grouped_by_bucket,
    bucket_display, bucket_order, make_plot, make_sample,
    normalize_rating, normalize_winding, rpm_grid, tidy_rows, display_rating_order,
)

st.set_page_config(page_title="주축 출력선도 생성기", page_icon="📈", layout="wide")
DATA_DIR = Path("data/motor_specs")
DATA_DIR.mkdir(parents=True, exist_ok=True)


SYSTEM_PROMPT = """
You are a technical transcription assistant for industrial motor PERFORMANCE TABLES.
The user uploads one or more screen captures of motor power/torque/RPM tables and charts.
Your job is ONLY to TRANSCRIBE values that are actually printed in the source.
Extract ALL winding sections separately (LOW / HIGH, if shown) and EACH duty curve
(S1 continuous, S2 n minutes, S3/S6 XX%, etc.) separately.

IMPORTANT RULES:
1) Prefer the NUMERIC specification tables rather than approximating numbers from plotted curves.
2) Do not fill in any missing/blank/dashed cell. Use null for absent power or torque.
3) Preserve every printed RPM anchor, including 0 rpm for zero power if 0 is printed.
4) At 0 rpm torque is NOT automatically 0. Use null if the catalog doesn't specify torque.
5) Preserve the distinction among 10%, 15%, 25%, S1, S2 and LOW/HIGH.
6) Ignore maximum current (A), winding resistance (mOhm) and power factor (percent)
   as curve points. Put such supplementary information in notes if relevant.
7) Do not calculate or infer the winding changeover speed if not explicitly stated.
8) If the image is only an Excel example without a motor numeric specification table,
   return curves=[] and a note that extraction is not possible.
9) If a number is unreadable, omit that cell (null) rather than guess.
10) Keep the motor_name empty if not displayed. maximum_rpm only if explicitly stated.
Return ONLY a JSON object following this exact example shape:
{
  "motor_name": "motor designation or empty string",
  "maximum_rpm": 20000,
  "curves": [
    {
      "winding": "LOW",
      "rating": "S1 Cont.",
      "points": [
        {"rpm": 0, "power_kw": 0, "torque_nm": null},
        {"rpm": 1415, "power_kw": 6.8, "torque_nm": 45.7}
      ]
    }
  ],
  "notes": []
}
The values above are an example of FORMAT ONLY. Never copy example values unless present in image.
Use JSON null for unknown numeric values. maximum_rpm should be null when not explicitly specified.
All numbers must be JSON numbers, not formatted strings. No prose, no markdown.
"""


def secret(name: str) -> str:
    try:
        return str(st.secrets.get(name, ""))
    except (FileNotFoundError, AttributeError, KeyError):
        return ""


def record_payload(rows, motor_name, rpm_limit, notes, record_id) -> dict:
    return {
        "record_id": record_id,
        "motor_name": motor_name,
        "max_rpm": rpm_limit,
        "notes": notes,
        "saved_at": datetime.now(timezone.utc).isoformat(),
        "rows": rows,
        "source_method": "motor-catalog-screenshot / manual-review",
        "schema_version": 1,
    }


def slug(s: str) -> str:
    s = re.sub(r"[^0-9a-zA-Z가-힣_-]+", "_", s.strip())
    return (s[:64] or "motor")


def local_path(payload: dict) -> Path:
    return DATA_DIR / f"{slug(payload['motor_name'])}_{payload['record_id']}.json"


def save_local(payload: dict):
    local_path(payload).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def github_config() -> tuple[str, str, str]:
    return secret("GITHUB_TOKEN"), secret("GITHUB_REPO"), secret("GITHUB_BRANCH") or "main"


def github_request(method, url, token, **kwargs):
    headers = {"Authorization": f"Bearer {token}",
               "Accept": "application/vnd.github+json",
               "X-GitHub-Api-Version": "2022-11-28"}
    return requests.request(method, url, headers=headers, timeout=25, **kwargs)


def github_save(payload: dict) -> str:
    token, repo, branch = github_config()
    if not token or not repo:
        return "GitHub 자동 저장 미설정 (현재 앱 서버 및 JSON 다운로드만 가능)"
    repo_info = github_request("GET", f"https://api.github.com/repos/{repo}", token)
    if repo_info.status_code != 200:
        raise RuntimeError(f"GitHub 저장소 접근 실패 HTTP {repo_info.status_code}")
    if not repo_info.json().get("private", False):
        raise RuntimeError("모터 사양의 공개 유출을 막기 위해 Public 저장소에는 자동 저장하지 않습니다. Private 저장소를 지정하십시오.")
    relpath = f"motor_specs/{slug(payload['motor_name'])}_{payload['record_id']}.json"
    url = f"https://api.github.com/repos/{repo}/contents/{relpath}"
    before = github_request("GET", url, token, params={"ref": branch})
    existing = before.json().get("sha") if before.status_code == 200 else None
    if before.status_code not in (200, 404):
        raise RuntimeError(f"GitHub 조회 실패 HTTP {before.status_code}: {before.text[:300]}")
    raw = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    body = {"message": f"Save motor spec {payload['motor_name']}",
            "content": base64.b64encode(raw).decode("ascii"), "branch": branch}
    if existing:
        body["sha"] = existing
    result = github_request("PUT", url, token, json=body)
    if result.status_code not in (200, 201):
        raise RuntimeError(f"GitHub 저장 실패 HTTP {result.status_code}: {result.text[:300]}")
    return f"GitHub 저장 완료: {repo}/{relpath}"


def github_records() -> list[dict]:
    token, repo, branch = github_config()
    if not token or not repo:
        return []
    url = f"https://api.github.com/repos/{repo}/contents/motor_specs"
    res = github_request("GET", url, token, params={"ref": branch})
    if res.status_code == 404:
        return []
    res.raise_for_status()
    return [item for item in res.json() if item.get("name", "").endswith(".json")]


def github_load(filename: str) -> dict:
    token, repo, branch = github_config()
    url = f"https://api.github.com/repos/{repo}/contents/motor_specs/{filename}"
    res = github_request("GET", url, token, params={"ref": branch})
    res.raise_for_status()
    data = res.json()
    return json.loads(base64.b64decode(data["content"]).decode("utf-8"))


def commit_save(rows, model_name, limit, notes, record_id) -> tuple[dict, str]:
    payload = record_payload(rows, model_name, limit, notes, record_id)
    save_local(payload)
    github_message = github_save(payload) if github_config()[0] and github_config()[1] else "GitHub 연결 없음"
    return payload, github_message


def extract_images(uploaded, api_key: str, model: str) -> dict:
    """Gemini REST API: no Pydantic or google-genai dependency required.

    The HTTP response is untrusted transcription data, not a guaranteed motor spec.
    It must be reviewed by a person before using it for design decisions.
    """
    if not re.fullmatch(r"[A-Za-z0-9_.-]+", model.strip()):
        raise ValueError("Gemini 모델명은 영문/숫자/하이픈/마침표/밑줄만 허용합니다.")

    parts = [{"text": SYSTEM_PROMPT}]
    raw_total = 0
    for f in uploaded:
        raw = f.getvalue()
        raw_total += len(raw)
        if len(raw) > 10 * 1024 * 1024:
            raise ValueError(f"{f.name}: 이미지 1장당 10MB 이하로 줄여주세요.")
        # Gemini inline-data has a request size limitation including base64 overhead.
        if raw_total > 12 * 1024 * 1024:
            raise ValueError("이미지 총 용량을 12MB 이하로 줄이거나 나누어 올려주세요.")
        mime = "image/png" if f.type == "image/png" else "image/jpeg"
        parts.append({"inline_data": {
            "mime_type": mime,
            "data": base64.b64encode(raw).decode("ascii"),
        }})
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model.strip()}:generateContent"
    payload = {
        "contents": [{"role": "user", "parts": parts}],
        "generationConfig": {"temperature": 0, "responseMimeType": "application/json"},
    }
    try:
        reply = requests.post(
            url, json=payload,
            headers={"x-goog-api-key": api_key, "Content-Type": "application/json"},
            timeout=120,
        )
    except requests.RequestException as exc:
        raise RuntimeError(f"Gemini API에 연결하지 못했습니다: {type(exc).__name__}") from exc
    if not reply.ok:
        # Avoid displaying API-key-bearing URLs or the full request/response body.
        try:
            info = reply.json().get("error", {})
            detail = str(info.get("message", ""))[:400]
        except ValueError:
            detail = ""
        raise RuntimeError(f"Gemini API HTTP {reply.status_code}: {detail}")
    try:
        data = reply.json()
        candidates = data.get("candidates", [])
        text_parts = candidates[0]["content"]["parts"]
        raw_json = "\n".join(part["text"] for part in text_parts if "text" in part).strip()
        # Some models may return a fenced JSON block even with JSON output mode.
        if raw_json.startswith("```"):
            raw_json = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw_json).strip()
        catalog = json.loads(raw_json)
    except (ValueError, KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("Gemini가 유효한 모터 사양 JSON을 반환하지 않았습니다. 이미지나 모델을 확인하세요.") from exc
    if not isinstance(catalog, dict) or not isinstance(catalog.get("curves"), list):
        raise RuntimeError("AI 응답에 curves 배열이 없습니다. 다른 이미지를 사용하거나 수동으로 입력하세요.")
    return catalog


def flatten(catalog: dict) -> list[dict]:
    """Normalize untrusted JSON records, letting tidy_rows reject bad numbers."""
    rows = []
    for curve in catalog.get("curves", []):
        if not isinstance(curve, dict):
            continue
        points = curve.get("points", [])
        if not isinstance(points, list):
            continue
        for p in points:
            if not isinstance(p, dict):
                continue
            rows.append({
                "권선": normalize_winding(curve.get("winding", "LOW")),
                "운전조건": normalize_rating(curve.get("rating", "S1")),
                "RPM": p.get("rpm"),
                "출력(kW)": p.get("power_kw"),
                "토크(N·m)": p.get("torque_nm"),
            })
    return rows


def use_record(payload: dict):
    st.session_state.rows = payload.get("rows", [])
    st.session_state.motor_name = payload.get("motor_name", "모터 모델명")
    st.session_state.max_rpm = max(100, int(payload.get("max_rpm") or 20000))
    st.session_state.notes = payload.get("notes", "")
    st.session_state.record_id = payload.get("record_id", str(uuid.uuid4())[:8])
    st.session_state.editor_version += 1
    st.session_state.pop("excel_bytes", None)
    st.session_state.pop("excel_key", None)


def new_id():
    return uuid.uuid4().hex[:10]


# 클립보드 이미지를 기존 파일 업로드와 동일하게 처리하기 위한 어댑터.
# 원본 이미지 바이트는 세션 내에서만 보관하며, 자동으로 파일로 저장하지 않습니다.
class ClipboardImage:
    def __init__(self, name: str, raw: bytes):
        self.name = name
        self.type = "image/png"
        self._raw = raw

    def getvalue(self) -> bytes:
        return self._raw


def clipboard_files() -> list[ClipboardImage]:
    return [ClipboardImage(item["name"], item["data"])
            for item in st.session_state.get("clipboard_images", [])]


def add_clipboard_image(image) -> tuple[bool, str]:
    """Clipboard -> PNG; exact same bytes flow as a normal uploaded file."""
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    raw = buffer.getvalue()
    if len(raw) > 10 * 1024 * 1024:
        return False, "캡처 이미지가 10MB를 초과합니다. 캡처 범위를 줄인 뒤 다시 등록하세요."
    digest = hashlib.sha256(raw).hexdigest()
    entries = st.session_state.setdefault("clipboard_images", [])
    if any(item["sha256"] == digest for item in entries):
        return False, "같은 캡처 이미지가 이미 등록되어 있습니다."
    filename = f"clipboard_{len(entries) + 1:02d}_{digest[:8]}.png"
    entries.append({"name": filename, "sha256": digest, "data": raw})
    return True, f"캡처 등록 완료: {filename}"


# ─────────────────────────────────────────────────────────
# 화면 구성 : 이미지 등록 → 인식 → 검토/저장 → 그래프/엑셀
# ─────────────────────────────────────────────────────────
def fingerprint_uploads(files) -> str | None:
    if not files:
        return None
    digest = hashlib.sha256()
    for f in files:
        digest.update(f.name.encode("utf-8", errors="replace"))
        digest.update(f.getvalue())
    return digest.hexdigest()


def file_audit(files):
    """파일을 선택했다는 사실과 실제 이미지를 분석했다는 사실은 구분한다."""
    infos, errors = [], []
    for f in files or []:
        raw = f.getvalue()
        png = raw.startswith(b"\x89PNG\r\n\x1a\n")
        jpeg = raw.startswith(b"\xff\xd8\xff")
        valid_image = png or jpeg
        if not valid_image:
            errors.append(f"{f.name}: PNG/JPEG 이미지 형식이 아닙니다.")
        if not raw:
            errors.append(f"{f.name}: 파일 내용이 비어 있습니다.")
        if len(raw) > 10 * 1024 * 1024:
            errors.append(f"{f.name}: AI 인식은 한 장당 10MB까지 지원합니다.")
        infos.append({"파일명": f.name, "크기": f"{len(raw) / (1024*1024):.2f} MB",
                      "이미지 확인": "확인" if valid_image else "형식 오류"})
    if sum(len(f.getvalue()) for f in files or []) > 12 * 1024 * 1024:
        errors.append("전체 이미지 용량이 12MB를 초과했습니다. 나누어서 등록하세요.")
    return infos, errors


def content_signature(rows, name: str, limit: int, notes: str) -> str:
    raw = json.dumps({"rows": rows, "name": name, "max": limit, "notes": notes},
                     ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def reset_excel():
    st.session_state.pop("excel_bytes", None)
    st.session_state.pop("excel_key", None)


for k, v in {
    "rows": [], "motor_name": "", "max_rpm": 20000,
    "record_id": new_id(), "notes": "", "editor_version": 0,
    "switch_rpm_user": 0, "show_switch_mode": True,
    "recognized_signature": None, "recognition_report": None,
    "recognition_error": "", "saved_signature": None,
    "saved_message": "", "data_origin": "미등록",
    "clipboard_images": [], "clipboard_feedback": "", "clipboard_button_version": 0,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

app_password = secret("APP_PASSWORD")
if app_password and not st.session_state.get("authenticated", False):
    st.title("🔐 주축 출력선도 생성기")
    attempt = st.text_input("앱 접속 비밀번호", type="password")
    if st.button("접속", type="primary"):
        if attempt == app_password:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    st.stop()

with st.sidebar:
    st.header("⚙️ 설정 및 저장 내역")
    with st.expander("AI 이미지 인식 설정", expanded=False):
        api_key = st.text_input("Gemini API 키", value=secret("GEMINI_API_KEY"),
                                type="password", help="Streamlit Secrets 설정을 권장합니다.")
        preset_models = [
            "gemini-3.8-flash",
            "gemini-2.5-flash-lite",
            "gemini-2.5-pro",
            "직접 입력",
        ]
        default_model = secret("GEMINI_MODEL") or "gemini-3.8-flash"
        default_index = preset_models.index(default_model) if default_model in preset_models else len(preset_models) - 1
        selected_model = st.selectbox("Google Gemini 모델 선택", preset_models, index=default_index,
                                      help="기본값은 gemini-3.8-flash 입니다. Google 정책 변경 시 다른 모델을 선택하거나 직접 입력할 수 있습니다.")
        if selected_model == "직접 입력":
            model = st.text_input("직접 입력 모델명", value=default_model if default_model not in preset_models else "gemini-3.8-flash")
        else:
            model = selected_model
        st.caption("기본 권장 모델: gemini-3.8-flash · 필요 시 다른 Google 모델을 직접 선택하거나 직접 입력할 수 있습니다. 인식 시 이미지가 Google Gemini API로 전달됩니다.")
    if not app_password:
        st.warning("APP_PASSWORD 미설정: 공개 앱이라면 접근 제한 설정을 권장합니다.")
    with st.expander("저장된 모터 사양 불러오기", expanded=False):
        paths = sorted(DATA_DIR.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
        if paths:
            chosen = st.selectbox("이 서버에 저장된 사양", paths, format_func=lambda f: f.stem)
            if st.button("서버 저장본 불러오기", use_container_width=True):
                use_record(json.loads(chosen.read_text(encoding="utf-8")))
                st.session_state.data_origin = "서버 JSON"
                st.session_state.recognized_signature = None
                st.session_state.recognition_report = None
                st.session_state.saved_signature = None
                st.session_state.saved_message = "불러온 데이터를 확인해 주세요."
                st.rerun()
        else:
            st.caption("현재 서버에 저장된 사양이 없습니다.")
        if github_config()[0] and github_config()[1]:
            try:
                remote = github_records()
                if remote:
                    name = st.selectbox("GitHub 영구 저장 사양", [i["name"] for i in remote])
                    if st.button("GitHub 저장본 불러오기", use_container_width=True):
                        use_record(github_load(name))
                        st.session_state.data_origin = "GitHub JSON"
                        st.session_state.recognized_signature = None
                        st.session_state.recognition_report = None
                        st.session_state.saved_signature = None
                        st.session_state.saved_message = "불러온 데이터를 확인해 주세요."
                        st.rerun()
            except Exception as exc:
                st.warning(f"GitHub 저장 목록을 읽을 수 없습니다: {exc}")
        else:
            st.caption("영구 저장을 원하면 별도 Private GitHub 저장소 연동이 필요합니다.")
    st.caption("서버 임시 저장은 재배포 후 사라질 수 있습니다. JSON 백업 또는 GitHub 저장을 이용하세요.")

st.title("⚙️ 주축 출력선도 생성기")
st.caption("모터 모델명 입력 → 사진 등록 → AI 숫자 인식 → 값 검토·저장 → 권선절환 반영 출력선도/엑셀 내보내기")

# 상단 고정 진행 상태. 업로드 파일 접수, AI 인식, 데이터 저장은 각각 별개의 상태입니다.
files_at_top = list(st.session_state.get('motor_catalog_images') or []) + clipboard_files()
current_upload_sig = fingerprint_uploads(files_at_top)
top_file_info, top_file_errors = file_audit(files_at_top)
top_uploaded = bool(files_at_top and not top_file_errors)
top_recognized = bool(top_uploaded and current_upload_sig == st.session_state.recognized_signature)
top_rows = bool(st.session_state.rows)
top_saved = bool(st.session_state.saved_signature)
with st.container(border=True):
    st.markdown('**진행 현황 — 사진만 선택한 상태와 AI 분석 완료 상태는 다릅니다.**')
    c1, c2, c3, c4 = st.columns(4)
    c1.metric('① 사진 등록', f'{len(files_at_top)}장 완료' if top_uploaded else '등록 대기')
    c2.metric('② AI 인식', '완료' if top_recognized else ('재인식 필요' if top_uploaded and top_rows and st.session_state.data_origin == 'AI 인식' else '실행 대기'))
    c3.metric('③ 사양 검토', f'{len(st.session_state.rows)}개 기준점' if top_rows else '검토 대기')
    c4.metric('④ 사양 저장', '저장 이력 있음' if top_saved else '저장 대기')
    if top_file_errors:
        st.error('선택한 사진 파일에 문제가 있습니다. ① 사진 등록 탭에서 세부 내용을 확인해 주세요.')
    elif top_uploaded:
        filename_line = ' · '.join(f.name for f in files_at_top)
        st.caption(f'현재 선택된 파일: {filename_line[:240]}' + ('…' if len(filename_line) > 240 else ''))
        if not top_recognized:
            st.info('**사진 등록까지만 완료**되었습니다. ①에서 「AI 인식 시작」을 눌러야 데이터가 추출됩니다.')
        else:
            st.success('현재 사진의 AI 인식 결과가 준비되었습니다. ②에서 수치를 확인하고 저장해 주세요.')

# 네 개의 작업 화면을 분리하여 필요한 정보에 집중할 수 있게 합니다.
step_upload, step_review, step_plot, step_excel = st.tabs([
    "① 사진 등록·AI 인식", "② 추출값 검토·저장", "③ 통합 출력선도", "④ 엑셀 다운로드"
])

with step_upload:
    st.subheader("① 모터 모델명 입력 및 사양 이미지 등록")
    st.caption("먼저 모터 모델명을 입력하세요. 그 다음 캡처 또는 사진 파일을 등록합니다. 권선절환 포인트를 입력하면 ③ 그래프와 ④ 엑셀 계산표가 즉시 같은 기준으로 갱신됩니다. 방법 A와 B를 함께 사용할 수 있습니다.")

    st.text_input("모터 모델명 (사진 등록 전에 먼저 입력)", key="motor_name", placeholder="예: Bill 1128/20000B")
    st.number_input("권선절환 포인트 RPM (입력 즉시 그래프/값 반영)", min_value=0, max_value=100000, step=100, key="switch_rpm_user", help="0이면 절환 미적용, 0보다 큰 값이면 LOW/HIGH 절환 적용 그래프와 값이 생성됩니다.")
    capture_col, upload_col = st.columns(2, gap="large")
    with capture_col:
        st.markdown("#### 📋 A. 화면 캡처 붙여넣기")
        st.markdown("**Windows:** `Win + Shift + S` → 모터 사양표 영역 드래그 → 아래 버튼 클릭")
        st.caption("클립보드에 복사된 이미지를 바로 받아옵니다. 캡처 파일을 저장할 필요가 없습니다.")
        paste_result = paste_image_button(
            "📋 복사한 캡처 이미지 등록",
            key=f"motor_clipboard_paste_button_{st.session_state.clipboard_button_version}",
            errors="raise",
        )
        if paste_result.image_data is not None:
            added, feedback = add_clipboard_image(paste_result.image_data)
            st.session_state.clipboard_feedback = feedback
            if added:
                st.rerun()
        if st.session_state.get("clipboard_feedback"):
            st.caption(st.session_state.clipboard_feedback)
        st.caption("Chrome 또는 Edge의 HTTPS 환경을 권장합니다. 브라우저가 요청하면 클립보드 읽기 권한을 허용하세요.")
        pasted_count = len(st.session_state.clipboard_images)
        if pasted_count:
            st.success(f"캡처 {pasted_count}장 등록됨")
            if st.button("캡처 등록 목록 비우기", key="clear_clipboard_images"):
                st.session_state.clipboard_images = []
                st.session_state.clipboard_button_version += 1  # Reset the component's stale paste result.
                st.session_state.clipboard_feedback = "캡처 이미지 목록을 비웠습니다."
                st.rerun()

    with upload_col:
        st.markdown("#### 📁 B. 사진 파일 업로드")
        st.caption("컴퓨터에 저장된 JPG/JPEG/PNG 파일을 선택하거나 끌어다 놓으세요.")
        uploads = st.file_uploader(
            "모터 카탈로그 사진 파일 선택",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key="motor_catalog_images",
            help="사진 등록 후 아래 AI 인식 버튼을 눌러야 수치가 분석됩니다.",
        )
        if uploads:
            st.success(f"파일 {len(uploads)}장 업로드됨")

    screenshots = list(uploads or []) + clipboard_files()
    st.divider()
    st.markdown("#### ✅ 실제 등록된 이미지 확인")
    st.caption("아래 미리보기에 표시된 이미지가 AI가 분석할 사진입니다. 파일 업로드와 캡처 등록을 합쳐 인식합니다.")
    upload_sig = fingerprint_uploads(screenshots)
    uploads_info, upload_errors = file_audit(screenshots)
    matched = bool(upload_sig and upload_sig == st.session_state.recognized_signature)
    if not screenshots:
        st.info("**사진 미등록** · 캡처 붙여넣기 버튼을 누르거나 사진 파일을 선택해 주세요.")
    else:
        if upload_errors:
            st.error("파일을 선택했지만 처리할 수 없는 이미지가 있습니다.")
            for error in upload_errors:
                st.warning(error)
        else:
            st.success(f"**사진 등록 확인: {len(screenshots)}장** · 파일을 정상적으로 받았습니다.")
            if matched:
                st.success("**AI 분석 완료** · 현재 선택된 사진의 수치를 추출했습니다. ② 사양 검토·저장 탭을 확인하세요.")
            else:
                st.info("**AI 분석 대기** · 사진 등록만 완료되었습니다. 아래 'AI 인식 시작' 버튼을 눌러야 숫자를 읽습니다.")
        st.dataframe(uploads_info, use_container_width=True, hide_index=True)
        with st.expander(f"📷 등록 사진 미리보기 ({len(screenshots)}장)", expanded=True):
            image_cols = st.columns(min(3, len(screenshots)))
            for i, file in enumerate(screenshots):
                with image_cols[i % len(image_cols)]:
                    raw = file.getvalue()
                    if raw.startswith(b"\x89PNG\r\n\x1a\n") or raw.startswith(b"\xff\xd8\xff"):
                        st.image(raw, caption=file.name, use_container_width=True)
                    else:
                        st.error(f"이미지 열기 실패: {file.name}")
        st.caption("※ 사진 파일 자체는 영구 저장되지 않습니다. 인식된 숫자는 ②에서 검토 후 저장할 수 있습니다.")

    if screenshots and not matched and st.session_state.rows and st.session_state.recognized_signature:
        st.warning("선택한 사진이 이전 분석 사진과 다릅니다. 아래 인식 버튼을 다시 눌러야 새 사진의 수치가 반영됩니다.")

    recognize_label = "🔄 다시 인식하기" if matched else "🔎 사진에서 모터 사양 자동 인식 시작"
    if st.button(recognize_label, type="primary", use_container_width=True,
                 disabled=not screenshots or bool(upload_errors)):
        if not api_key:
            st.session_state.recognition_error = "Gemini API 키가 없습니다. 왼쪽 'AI 이미지 인식 설정' 또는 Streamlit Secrets에 GEMINI_API_KEY를 지정해 주세요."
        else:
            st.session_state.recognition_error = ""
            try:
                with st.status("이미지를 분석하고 있습니다…", expanded=True) as progress:
                    st.write(f"1/3 · 업로드한 사진 {len(screenshots)}장 확인")
                    st.write("2/3 · Gemini AI에서 RPM·출력·토크 표 추출")
                    catalog = extract_images(screenshots, api_key, model)
                    parsed, issues = tidy_rows(flatten(catalog))
                    if not parsed:
                        raise ValueError("인식 가능한 RPM / 출력 / 토크 값이 없습니다. 숫자가 보이는 표 사진을 사용해 주세요.")
                    st.write(f"3/3 · {len(parsed)}개 기준점 · {len(group_rows(parsed))}개 운전곡선 분류 완료")
                    progress.update(label="AI 사양 인식 완료", state="complete", expanded=False)
                st.session_state.rows = parsed
                detected_name = catalog.get("motor_name") or ""
                if not st.session_state.motor_name:
                    st.session_state.motor_name = detected_name
                st.session_state.max_rpm = max(100, int(catalog.get("maximum_rpm") or max(x["RPM"] for x in parsed)))
                st.session_state.notes = " / ".join(map(str, catalog.get("notes", [])))
                st.session_state.record_id = new_id()
                st.session_state.editor_version += 1
                st.session_state.recognized_signature = upload_sig
                st.session_state.recognition_report = {
                    "files": [f.name for f in screenshots],
                    "points": len(parsed),
                    "curves": len(group_rows(parsed)),
                    "warnings": issues,
                    "at": datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S"),
                }
                st.session_state.saved_signature = None
                st.session_state.saved_message = ""
                st.session_state.data_origin = "AI 인식"
                reset_excel()
                st.rerun()
            except Exception as exc:
                st.session_state.recognition_error = str(exc)
    if st.session_state.recognition_error:
        st.error(f"AI 인식 실패: {st.session_state.recognition_error}")
    report = st.session_state.recognition_report
    if report and matched:
        st.markdown("##### 인식 결과 요약")
        k1, k2, k3 = st.columns(3)
        k1.metric("인식 기준점", f"{report['points']}개")
        k2.metric("운전 곡선", f"{report['curves']}개")
        k3.metric("등록 사진", f"{len(report['files'])}장")
        st.caption(f"마지막 인식: {report['at']} · AI 인식 완료는 수치 검토나 저장 완료를 의미하지 않습니다.")
        st.markdown('**사진에서 추출한 운전조건 (빠른 확인용)**')
        extracted = group_rows(st.session_state.rows)
        check_rows = []
        for key in sorted(extracted, key=lambda k: (k[0] != 'LOW', k[0] != 'HIGH', display_rating_order(k[1]))):
            pts = extracted[key]
            p_values = [v['출력(kW)'] for v in pts if v['출력(kW)'] is not None]
            t_values = [v['토크(N·m)'] for v in pts if v['토크(N·m)'] is not None]
            check_rows.append({'권선': key[0], '운전조건': key[1], 'RPM 범위': f"{pts[0]['RPM']:g}~{pts[-1]['RPM']:g}",
                               '기준점': len(pts), '최대 출력(kW)': max(p_values) if p_values else None,
                               '최대 토크(N·m)': max(t_values) if t_values else None})
        st.dataframe(check_rows, hide_index=True, use_container_width=True)
        st.warning('AI가 인식한 값은 오독 가능성이 있습니다. 아래 ② 탭에서 원본 사진과 비교해 검토·저장해야 합니다.')
        if report["warnings"]:
            with st.expander(f"AI 수치 확인 필요 ({len(report['warnings'])}건)"):
                for item in report["warnings"][:25]:
                    st.warning(item)
        if st.button("② 추출된 데이터 검토하러 가기", use_container_width=True):
            st.info("바로 위의 **② 사양 검토·저장** 탭을 클릭하세요.")
    st.divider()
    shortcut1, shortcut2 = st.columns(2)
    with shortcut1:
        if st.button("예제 데이터로 기능 확인", use_container_width=True):
            sample, note = make_sample()
            st.session_state.rows = sample
            st.session_state.motor_name = "LOW-HIGH 모터 예제"
            st.session_state.switch_rpm_user = 4500
            st.session_state.max_rpm = 20000
            st.session_state.notes = note
            st.session_state.record_id = new_id()
            st.session_state.editor_version += 1
            st.session_state.recognized_signature = None
            st.session_state.recognition_report = None
            st.session_state.recognition_error = ""
            st.session_state.saved_signature = None
            st.session_state.data_origin = "예제"
            reset_excel()
            st.rerun()
    with shortcut2:
        if st.button("빈 입력표로 시작", use_container_width=True):
            st.session_state.rows = []
            st.session_state.motor_name = ""
            st.session_state.switch_rpm_user = 0
            st.session_state.max_rpm = 20000
            st.session_state.notes = ""
            st.session_state.record_id = new_id()
            st.session_state.editor_version += 1
            st.session_state.recognized_signature = None
            st.session_state.recognition_report = None
            st.session_state.recognition_error = ""
            st.session_state.saved_signature = None
            st.session_state.data_origin = "수동 입력"
            reset_excel()
            st.rerun()

with step_review:
    st.subheader("② 인식된 모터 사양 확인 및 저장")
    if screenshots:
        with st.expander(f"📷 인식에 사용한 원본 이미지 다시 확인 ({len(screenshots)}장)", expanded=False):
            st.caption("①에 등록된 원본과 아래 인식값을 비교하세요.")
            reference_cols = st.columns(min(2, len(screenshots)))
            for idx, reference_image in enumerate(screenshots):
                with reference_cols[idx % len(reference_cols)]:
                    image_bytes = reference_image.getvalue()
                    if image_bytes.startswith(b"\x89PNG\r\n\x1a\n") or image_bytes.startswith(b"\xff\xd8\xff"):
                        st.image(image_bytes, caption=reference_image.name, use_container_width=True)
    if st.session_state.data_origin == "AI 인식":
        if matched:
            st.success("현재 사진의 AI 인식 데이터입니다. 원본 사진과 비교 후 저장하세요.")
        elif screenshots:
            st.warning("현재 등록된 사진과 표시된 인식값이 다릅니다. ①에서 다시 인식하세요.")
        else:
            st.info("이전 AI 인식 데이터입니다. 필요하면 새 사진을 등록하여 다시 인식하세요.")
    elif st.session_state.rows:
        st.info(f"현재 데이터 출처: {st.session_state.data_origin}")
    else:
        st.info("① 사진 등록·인식에서 사진을 분석하거나 예제/수동 입력을 선택하세요.")

    a, b = st.columns([1.25, 0.75])
    with a:
        model_name = st.text_input("모터 모델명", value=st.session_state.motor_name)
    with b:
        rpm_limit = st.number_input("최대 RPM (그래프/계산표)", min_value=100,
                                    max_value=100000, value=int(st.session_state.max_rpm), step=100)
    remark = st.text_area("기타 사양 및 참고사항", value=st.session_state.notes, height=75)
    imported = st.file_uploader("저장해 둔 모터 사양 JSON 불러오기 (선택)", type=["json"], key="json_import")
    if st.button("JSON 데이터 불러오기", disabled=imported is None):
        try:
            payload = json.loads(imported.getvalue().decode("utf-8-sig"))
            use_record(payload)
            st.session_state.recognized_signature = None
            st.session_state.recognition_report = None
            st.session_state.saved_signature = None
            st.session_state.data_origin = "JSON 불러오기"
            st.rerun()
        except Exception as exc:
            st.error(f"JSON 파일을 읽지 못했습니다: {exc}")

    editable = st.session_state.rows or [
        {"권선": "LOW", "운전조건": "S1 Cont.", "RPM": 0, "출력(kW)": 0.0, "토크(N·m)": None},
        {"권선": "LOW", "운전조건": "S1 Cont.", "RPM": 1500, "출력(kW)": None, "토크(N·m)": None},
    ]
    st.caption("아래는 AI가 읽은 값입니다. 반드시 원본과 대조한 후 수정·저장하세요. 빈칸은 미확인 수치입니다.")
    updated = st.data_editor(
        editable, hide_index=True, num_rows="dynamic", use_container_width=True,
        key=f"motor_editor_{st.session_state.editor_version}",
        column_config={
            "권선": st.column_config.TextColumn("권선 LOW/HIGH"),
            "운전조건": st.column_config.TextColumn("운전조건 S1/S2/S3"),
            "RPM": st.column_config.NumberColumn("RPM", min_value=0, format="%.1f"),
            "출력(kW)": st.column_config.NumberColumn("출력 (kW)", min_value=0, format="%.3f"),
            "토크(N·m)": st.column_config.NumberColumn("토크 (N·m)", min_value=0, format="%.3f"),
        },
    )
    valid, warnings = tidy_rows(updated)
    groups = group_rows(valid)
    if valid:
        st.markdown("##### 추출 조건별 검토 요약")
        summary = []
        for (winding, rating), points in sorted(groups.items(), key=lambda item: (item[0][0] != 'LOW', item[0][0] != 'HIGH', display_rating_order(item[0][1]))):
            kw_values = [v["출력(kW)"] for v in points if v["출력(kW)"] is not None]
            nm_values = [v["토크(N·m)"] for v in points if v["토크(N·m)"] is not None]
            summary.append({"권선": winding, "운전조건": rating, "기준점 수": len(points),
                            "최소 RPM": points[0]["RPM"], "최대 RPM": points[-1]["RPM"],
                            "최대 출력(kW)": max(kw_values) if kw_values else None,
                            "최대 토크(N·m)": max(nm_values) if nm_values else None})
        st.dataframe(summary, hide_index=True, use_container_width=True)
        st.caption(f"인식/입력값 {len(valid)}개 기준점 · {len(groups)}개 곡선")
    if warnings:
        with st.expander(f"⚠️ 검토할 데이터 ({len(warnings)}건)", expanded=True):
            for warning in warnings[:30]:
                st.warning(warning)
    pending_sig = content_signature(valid, model_name, int(rpm_limit), remark) if valid else None
    saved_current = bool(pending_sig and st.session_state.saved_signature == pending_sig)
    stale_photo = bool(screenshots and st.session_state.recognized_signature
                       and not matched and st.session_state.data_origin == "AI 인식")
    if saved_current:
        st.success(f"✅ **수치 저장 완료** · {st.session_state.saved_message}")
    else:
        st.info("**검토·저장 대기** · 아래 버튼을 눌러야 수정한 수치가 저장됩니다.")
    save_col, json_col = st.columns(2)
    with save_col:
        if st.button("✅ 검토한 모터 사양 저장", type="primary", use_container_width=True,
                     disabled=not valid or stale_photo):
            try:
                st.session_state.rows = valid
                st.session_state.motor_name = model_name
                st.session_state.max_rpm = int(rpm_limit)
                st.session_state.notes = remark
                st.session_state.editor_version += 1
                payload, sync = commit_save(valid, model_name, int(rpm_limit), remark, st.session_state.record_id)
                st.session_state.saved_signature = content_signature(valid, model_name, int(rpm_limit), remark)
                st.session_state.saved_message = "서버에 JSON 저장됨 · " + sync
                reset_excel()
                st.rerun()
            except Exception as exc:
                st.error(f"저장 실패: {exc}")
    with json_col:
        current_payload = record_payload(valid, model_name, int(rpm_limit), remark, st.session_state.record_id)
        st.download_button("📥 모터 사양 JSON 백업", data=json.dumps(current_payload, ensure_ascii=False, indent=2),
                           file_name=f"{slug(model_name)}_motor_spec.json", mime="application/json",
                           use_container_width=True, disabled=not valid)
    if stale_photo:
        st.warning("사진이 변경되어 저장을 잠시 막았습니다. ①에서 새 사진을 인식하거나 빈 입력표/예제로 시작하세요.")

with step_plot:
    st.subheader("③ 출력·토크 통합 선도")
    if not valid:
        st.info("먼저 ①에서 사진을 등록하고 ②에서 데이터를 확인하세요.")
    else:
        if not saved_current:
            st.warning("그래프는 현재 검토표의 수치로 그려집니다. 저장이 아직 완료되지 않았다면 ②에서 저장하세요.")
        show_mode = st.radio("표시할 항목", ["출력 + 토크", "출력만", "토크만"], horizontal=True, index=0)
        switch_enabled = st.checkbox("권선절환 포인트 적용", value=bool(st.session_state.switch_rpm_user > 0))
        switch_rpm = st.number_input("권선절환 RPM", min_value=0, max_value=int(rpm_limit),
                                     value=min(int(st.session_state.switch_rpm_user), int(rpm_limit)), step=100)
        st.session_state.switch_rpm_user = int(switch_rpm) if switch_enabled else 0
        apply_switch = bool(switch_enabled and switch_rpm > 0)
        show_raw_points = st.checkbox("원본 기준점 점 표시", value=True)

        fig = make_plot(valid, max_rpm=int(rpm_limit),
                        show_power=show_mode != "토크만",
                        show_torque=show_mode != "출력만",
                        switch_rpm=int(switch_rpm) if apply_switch else None,
                        apply_switch=apply_switch,
                        show_raw_points=show_raw_points)
        st.plotly_chart(fig, use_container_width=True)
        if apply_switch:
            st.caption("예시 이미지처럼 권선절환 RPM을 기준으로 LOW/HIGH 곡선을 자동 결합해 표시합니다. LOW 계열은 회색, HIGH 출력은 파란색, HIGH 토크는 빨간색으로 표시됩니다.")
        else:
            st.caption("권선절환 포인트가 없으면 LOW/HIGH 원본 곡선을 그대로 표시합니다. 색상은 예시 이미지 느낌에 맞춰 LOW는 회색, HIGH 출력은 파란색, HIGH 토크는 빨간색 계열입니다.")
        st.download_button("그래프 HTML 다운로드", data=fig.to_html(include_plotlyjs=True, full_html=True),
                           file_name=f"{slug(model_name)}_diagram.html", mime="text/html")

        st.markdown("##### 권선절환 적용값 미리보기")
        grouped_bucket = grouped_by_bucket(valid)
        buckets = sorted({b for _, b in grouped_bucket}, key=bucket_order)
        preview_rows = []
        preview_grid = rpm_grid(valid, limit=int(rpm_limit), step=500, switch_rpm=int(switch_rpm) if apply_switch else None)[:80]
        for rpm in preview_grid:
            row = {"RPM": rpm}
            for bucket in buckets:
                label = bucket_display(bucket)
                if apply_switch:
                    kw, nm, src = calc_switched_at(grouped_bucket, bucket, rpm, int(switch_rpm))
                    row[f"{label} 출력(kW)"] = round(kw, 3) if kw is not None else None
                    row[f"{label} 토크(N·m)"] = round(nm, 3) if nm is not None else None
                    row[f"{label} 적용권선"] = src
                else:
                    for winding in ["LOW", "HIGH"]:
                        pts = grouped_bucket.get((winding, bucket))
                        if pts:
                            kw, nm = calc_at(pts, rpm)
                            row[f"{winding} {label} 출력(kW)"] = round(kw, 3) if kw is not None else None
                            row[f"{winding} {label} 토크(N·m)"] = round(nm, 3) if nm is not None else None
            preview_rows.append(row)
        st.dataframe(preview_rows, hide_index=True, use_container_width=True)

with step_excel:
    st.subheader("④ RPM별 계산표 및 엑셀 다운로드")
    if not valid:
        st.info("먼저 데이터를 입력한 후 엑셀을 생성할 수 있습니다.")
    else:
        st.caption("엑셀 시트: 통합 선도 / RPM 계산표 / 권선절환 적용값 / 기준점 입력 / 사용 안내")
        st.success("**RPM 계산표 배열:** LOW 출력(S3 10%→15%→25%→S2→S1) → LOW 토크(같은 순서) → HIGH 출력 → HIGH 토크. 카탈로그에 없는 조건은 추가하지 않습니다.")
        current_switch = int(st.session_state.switch_rpm_user or 0)
        if current_switch > 0:
            st.info(f"현재 권선절환 포인트: {current_switch} rpm · 엑셀의 [권선절환 적용값] 시트와 [통합 선도]에 즉시 반영됩니다.")
        else:
            st.info("권선절환 포인트가 0이므로 LOW/HIGH 원본 계산표 기준으로 엑셀이 생성됩니다.")
        rpm_step = st.selectbox("RPM 계산 간격", [10, 25, 50, 100, 200, 500], index=3)
        stamp = content_signature(valid, model_name, int(rpm_limit), remark) + f"|{rpm_step}|{current_switch}"
        if st.session_state.get("excel_key") != stamp:
            st.session_state.pop("excel_bytes", None)
        if st.button("📊 엑셀 계산표 생성", type="primary", use_container_width=True):
            try:
                with st.spinner("엑셀 차트와 RPM 계산표를 생성하는 중입니다…"):
                    st.session_state.excel_bytes = generate_excel(
                        valid, model_name, max_rpm=int(rpm_limit), step=int(rpm_step),
                        switch_rpm=current_switch if current_switch > 0 else None,
                        notes=remark,
                    )
                    st.session_state.excel_key = stamp
                st.success("엑셀 생성 완료 · 아래 다운로드 버튼으로 파일을 저장하세요.")
            except Exception as exc:
                st.error(f"엑셀 생성 실패: {exc}")
        if st.session_state.get("excel_bytes"):
            st.download_button("📥 주축 출력선도 엑셀 다운로드 (.xlsx)",
                               data=st.session_state.excel_bytes,
                               file_name=f"{slug(model_name)}_spindle_diagram.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="primary", use_container_width=True)
        with st.expander("RPM 계산값 미리보기", expanded=False):
            preview = []
            winding_names = sorted({k[0] for k in groups}, key=lambda w: (w != 'LOW', w != 'HIGH', w))
            for rpm in rpm_grid(valid, limit=int(rpm_limit), step=int(rpm_step), switch_rpm=current_switch if current_switch > 0 else None)[:200]:
                line = {'RPM': rpm}
                for winding in winding_names:
                    keys = sorted((k for k in groups if k[0] == winding), key=lambda k: display_rating_order(k[1]))
                    for quantity in ('출력(kW)', '토크(N·m)'):
                        for key in keys:
                            kw, nm = calc_at(groups[key], rpm)
                            value = kw if quantity == '출력(kW)' else nm
                            line[f'{winding} / {quantity} / {key[1]}'] = round(value, 3) if value is not None else None
                preview.append(line)
            st.dataframe(preview, hide_index=True, use_container_width=True)

