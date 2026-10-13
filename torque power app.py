"""Streamlit web app: Screenshot -> motor specs -> diagram -> formula-based Excel.
Run: streamlit run "torque power app.py"
"""
from __future__ import annotations

import base64
import json
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path

import requests
import streamlit as st

from motor_core import (
    calc_at, generate_excel, group_rows, make_plot, make_sample,
    normalize_rating, normalize_winding, rpm_grid, tidy_rows,
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


for k, v in {
    "rows": [], "motor_name": "모터 모델명", "max_rpm": 20000,
    "record_id": new_id(), "notes": "", "editor_version": 0,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

app_password = secret("APP_PASSWORD")
if app_password:
    if not st.session_state.get("authenticated", False):
        st.title("🔐 주축 출력선도 생성기")
        attempt = st.text_input("앱 접속 비밀번호", type="password")
        if st.button("접속"):
            if attempt == app_password:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")
        st.stop()
else:
    st.warning("보안 안내: 이 앱은 현재 비밀번호가 설정되지 않았습니다. 공개 Streamlit 배포에서는 APP_PASSWORD를 Secrets에 설정하십시오. 업로드 자료가 제3자에게 공개되지 않도록 주의하십시오.")

st.title("주축 출력선도 자동 생성기")
st.caption("모터 카탈로그 캡처 → AI 표 수치 추출 → 검토·수정 → 출력/토크 통합 그래프 → 수식 포함 엑셀")

with st.sidebar:
    st.header("AI 이미지 인식")
    api_key = st.text_input("Gemini API 키", value=secret("GEMINI_API_KEY"), type="password",
                             help="Streamlit Cloud Secrets에 GEMINI_API_KEY를 설정하면 입력하지 않아도 됩니다.")
    model = st.text_input("Gemini 모델", value=secret("GEMINI_MODEL") or "gemini-2.5-flash")
    st.caption("AI 인식을 실행하면 업로드한 이미지가 Google Gemini API로 전달됩니다. 요금과 한도는 사용자의 Google API 계정 정책을 따릅니다.")
    st.divider()
    st.header("저장된 모터 사양")
    paths = sorted(DATA_DIR.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
    if paths:
        selected = st.selectbox("현재 앱 서버에 저장된 파일", paths, format_func=lambda p: p.stem)
        if st.button("선택 파일 불러오기", use_container_width=True):
            use_record(json.loads(selected.read_text(encoding="utf-8")))
            st.rerun()
    if github_config()[0] and github_config()[1]:
        try:
            remote = github_records()
            if remote:
                remote_name = st.selectbox("GitHub에 영구 저장된 파일", [x["name"] for x in remote])
                if st.button("GitHub 파일 불러오기", use_container_width=True):
                    use_record(github_load(remote_name))
                    st.rerun()
        except Exception as exc:
            st.warning(f"GitHub 목록 조회 실패: {exc}")
    else:
        st.info("GitHub 영구 저장을 사용하려면 배포 환경 Secrets에 GITHUB_TOKEN, GITHUB_REPO를 설정하십시오.")
    st.caption("주의: Streamlit Cloud의 로컬 파일 저장은 재시작·재배포 시 유실될 수 있습니다. GitHub 연동 또는 JSON 다운로드를 권장합니다.")

upload_col, data_col = st.columns([1.15, 1.0], gap="large")
with upload_col:
    st.subheader("① 카탈로그 이미지 등록")
    screenshots = st.file_uploader("LOW/HIGH 권선 표가 있는 캡처를 등록하십시오 (여러 장 가능)",
                                   type=["png", "jpg", "jpeg"], accept_multiple_files=True)
    if screenshots:
        with st.expander(f"업로드 이미지 미리보기 ({len(screenshots)}장)", expanded=True):
            for file in screenshots:
                st.image(file, caption=file.name, use_container_width=True)
    if st.button("이미지 자동 인식 및 사양 저장", type="primary", disabled=not bool(screenshots), use_container_width=True):
        if not api_key:
            st.error("Gemini API 키가 필요합니다. 왼쪽 사이드바에 입력하거나 Streamlit Secrets에 GEMINI_API_KEY를 설정하십시오.")
        else:
            try:
                with st.spinner("표의 수치와 운전 조건을 추출하고 있습니다..."):
                    catalog = extract_images(screenshots, api_key, model)
                    parsed, problems = tidy_rows(flatten(catalog))
                if not parsed:
                    st.error("읽을 수 있는 모터 RPM/출력/토크 표가 없습니다. 더 선명한 캡처 또는 직접 입력을 사용하십시오.")
                    if catalog.get("notes"):
                        st.info(" / ".join(map(str, catalog.get("notes", []))))
                else:
                    st.session_state.rows = parsed
                    st.session_state.motor_name = catalog.get("motor_name") or st.session_state.motor_name
                    st.session_state.max_rpm = max(100, int(catalog.get("maximum_rpm") or max(x["RPM"] for x in parsed)))
                    st.session_state.notes = " / ".join(map(str, catalog.get("notes", [])))
                    st.session_state.record_id = new_id()
                    st.session_state.editor_version += 1
                    try:
                        _, sync = commit_save(parsed, st.session_state.motor_name,
                                              st.session_state.max_rpm, st.session_state.notes,
                                              st.session_state.record_id)
                    except Exception as sync_exc:
                        sync = f"로컬 저장 성공, GitHub 동기화 실패: {sync_exc}"
                    st.session_state.pop("excel_bytes", None)
                    st.success(f"{len(parsed)}개 원본 기준점 추출 완료. {sync}")
                    if problems:
                        st.warning(" / ".join(problems[:8]))
            except Exception as exc:
                st.error(f"자동 인식 실패: {exc}")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("예제 모터 불러오기", use_container_width=True):
            sample, note = make_sample()
            st.session_state.rows = sample
            st.session_state.motor_name = "LOW-HIGH 모터 예제"
            st.session_state.max_rpm = 20000
            st.session_state.notes = note
            st.session_state.record_id = new_id()
            st.session_state.editor_version += 1
            st.session_state.pop("excel_bytes", None)
            st.rerun()
    with c2:
        if st.button("빈 입력표로 새로 시작", use_container_width=True):
            st.session_state.rows = []
            st.session_state.motor_name = "모터 모델명"
            st.session_state.max_rpm = 20000
            st.session_state.record_id = new_id()
            st.session_state.notes = ""
            st.session_state.editor_version += 1
            st.session_state.pop("excel_bytes", None)
            st.rerun()

with data_col:
    st.subheader("② 모터 모델 및 데이터 불러오기")
    model_name = st.text_input("모터 모델명", value=st.session_state.motor_name)
    rpm_limit = st.number_input("그래프 / RPM 계산표의 최대 회전수", min_value=100,
                               max_value=100000, value=int(st.session_state.max_rpm), step=100)
    remark = st.text_area("참고 사항", value=st.session_state.notes, height=90)
    input_json = st.file_uploader("이전에 내려받은 모터 사양 JSON 불러오기", type=["json"], key="json_import")
    if st.button("JSON 파일 적용", disabled=input_json is None, use_container_width=True):
        try:
            payload = json.loads(input_json.getvalue().decode("utf-8-sig"))
            use_record(payload)
            st.rerun()
        except Exception as exc:
            st.error(f"JSON 형식 오류: {exc}")

st.divider()
st.subheader("③ 인식된 사양 검토 및 직접 수정")
st.caption("표에서 오타와 누락된 수치를 반드시 확인하십시오. 행을 추가해 직접 입력할 수도 있습니다. 공란은 해당 값이 제공되지 않았다는 뜻입니다.")
editable = st.session_state.rows or [
    {"권선": "LOW", "운전조건": "S1 Cont.", "RPM": 0, "출력(kW)": 0.0, "토크(N·m)": None},
    {"권선": "LOW", "운전조건": "S1 Cont.", "RPM": 1500, "출력(kW)": None, "토크(N·m)": None},
]
updated = st.data_editor(
    editable, hide_index=True, num_rows="dynamic", use_container_width=True,
    key=f"motor_editor_{st.session_state.editor_version}",
    column_config={
        "권선": st.column_config.TextColumn("권선 LOW/HIGH", help="LOW, HIGH 등"),
        "운전조건": st.column_config.TextColumn("운전조건", help="예: S1 Cont., S2 15min, S3/S6 25%"),
        "RPM": st.column_config.NumberColumn("RPM", min_value=0, format="%.1f"),
        "출력(kW)": st.column_config.NumberColumn("출력 (kW)", min_value=0, format="%.3f"),
        "토크(N·m)": st.column_config.NumberColumn("토크 (N·m)", min_value=0, format="%.3f"),
    },
)
valid, warnings = tidy_rows(updated)
if warnings:
    with st.expander(f"입력 데이터 검증 경고 ({len(warnings)}건)", expanded=True):
        for item in warnings[:30]:
            st.warning(item)
if not valid:
    st.info("유효한 모터 데이터를 넣으면 그래프와 엑셀을 생성할 수 있습니다.")
    st.stop()

save_c, json_c = st.columns([1, 1])
with save_c:
    if st.button("검토한 모터 사양 확정 및 저장", type="primary", use_container_width=True):
        try:
            st.session_state.rows = valid
            st.session_state.motor_name = model_name
            st.session_state.max_rpm = int(rpm_limit)
            st.session_state.notes = remark
            st.session_state.editor_version += 1
            st.session_state.pop("excel_bytes", None)
            _, msg = commit_save(valid, model_name, int(rpm_limit), remark, st.session_state.record_id)
            st.success(f"저장 완료. {msg}")
            st.rerun()
        except Exception as exc:
            st.error(f"저장 오류: {exc}")
with json_c:
    payload_json = record_payload(valid, model_name, int(rpm_limit), remark, st.session_state.record_id)
    st.download_button("모터 사양 JSON 백업", data=json.dumps(payload_json, ensure_ascii=False, indent=2),
                       file_name=f"{slug(model_name)}_motor_spec.json", mime="application/json", use_container_width=True)

st.divider()
st.subheader("④ 출력·토크 통합 선도")
groups = group_rows(valid)
left, right = st.columns([1, 1])
with left:
    show_power = st.checkbox("출력(kW) 표시 · 실선", value=True)
    show_torque = st.checkbox("토크(N·m) 표시 · 점선", value=True)
with right:
    show_change = st.checkbox("권선 절환 RPM 표시 (참고용)", value=False)
    switch_rpm = st.number_input("권선 절환 RPM", min_value=0, max_value=int(rpm_limit), value=0, step=100,
                                 disabled=not show_change)
if show_change:
    st.info("절환 위치는 사용자 지정 참고선입니다. 원본에 없는 절환 RPM을 자동 추정하지 않습니다.")
if show_power or show_torque:
    fig = make_plot(valid, max_rpm=int(rpm_limit), show_power=show_power,
                    show_torque=show_torque, switch_rpm=switch_rpm if show_change else None)
    st.plotly_chart(fig, use_container_width=True)
    st.download_button("그래프 HTML 다운로드", data=fig.to_html(include_plotlyjs=True, full_html=True),
                       file_name=f"{slug(model_name)}_diagram.html", mime="text/html")
else:
    st.warning("출력 또는 토크 표시를 선택하십시오.")

st.divider()
st.subheader("⑤ 회전수별 엑셀 계산표 및 Excel 내장 그래프")
a, b = st.columns([1, 1])
with a:
    rpm_step = st.selectbox("RPM 계산 간격", [10, 25, 50, 100, 200, 500], index=3)
with b:
    st.metric("읽어들인 기준점", f"{len(valid)}개", f"{len(groups)}개 곡선", delta_color="off")
st.caption("계산 기준: 출력 선형 보간 후 토크 = 출력(kW) × 9549.2965855 / rpm. 자료 없는 구간은 외삽하지 않습니다. 엑셀 원본 기준점 수정 시 수식이 다시 계산됩니다.")
preview_rpm = rpm_grid(valid, limit=int(rpm_limit), step=int(rpm_step))
preview = []
for rpm in preview_rpm[:300]:
    line = {"RPM": rpm}
    for key, points in groups.items():
        kw, nm = calc_at(points, rpm)
        line[f"{key[0]}/{key[1]} 출력(kW)"] = round(kw, 3) if kw is not None else None
        line[f"{key[0]}/{key[1]} 토크(N·m)"] = round(nm, 3) if nm is not None else None
    preview.append(line)
with st.expander("RPM 계산표 미리보기 (상위 300행)"):
    st.dataframe(preview, hide_index=True, use_container_width=True)

stamp = f"{model_name}|{int(rpm_limit)}|{rpm_step}|{switch_rpm if show_change else ''}|{json.dumps(valid, ensure_ascii=False, sort_keys=True)}"
if st.session_state.get("excel_key") != stamp:
    st.session_state.pop("excel_bytes", None)
if st.button("엑셀 생성하기 (계산값 + 원본 입력 + 선도)", type="primary"):
    try:
        with st.spinner("엑셀 계산표와 차트를 작성하고 있습니다..."):
            st.session_state.excel_bytes = generate_excel(
                valid, model_name, max_rpm=int(rpm_limit), step=int(rpm_step),
                switch_rpm=int(switch_rpm) if show_change else None, notes=remark)
            st.session_state.excel_key = stamp
        st.success("엑셀 생성 완료: [통합 선도], [RPM 계산표], [기준점 입력], [사용 안내] 시트 포함")
    except Exception as exc:
        st.error(f"엑셀 생성 실패: {exc}")
if st.session_state.get("excel_bytes"):
    st.download_button("📥 모터 주축 출력선도.xlsx 다운로드", data=st.session_state.excel_bytes,
                       file_name=f"{slug(model_name)}_spindle_diagram.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

st.caption("참고: 이 앱은 카탈로그 전사 및 도표 작성을 지원합니다. 최종 설계·선정 자료에는 제조사 원본과 대조 검증이 필요합니다.")
