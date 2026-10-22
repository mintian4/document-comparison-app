from __future__ import annotations

import difflib
import html
import os
import re
import struct
import zipfile
import zlib
from io import BytesIO
from xml.etree import ElementTree

import streamlit as st


st.set_page_config(page_title="문서 변경 비교기", page_icon="📝", layout="wide")
st.title("📝 문서 변경 비교기")
st.write("이전 문서와 수정 문서를 비교해 추가·삭제된 문장을 표시합니다.")
st.caption("연습용 앱입니다. 실제 GMP/품질 문서의 공식 검토나 승인 절차를 대신하지 않습니다.")
st.warning("이 앱이 인터넷에 배포된 경우, 업로드한 문서는 호스팅 서버로 전송되어 처리됩니다. 공개 가능한 연습용 파일만 사용하고 민감한 문서는 올리지 마세요.")

SAMPLE_OLD = """문서번호: GMP-TR-001\n제목: 정제수 설비 일일 점검 절차\n\n1. 목적\n정제수 설비의 운전 상태를 매일 확인한다.\n\n2. 점검 항목\n- 저장탱크 온도: 20~25°C\n- 순환 압력: 1.0~1.5 bar\n- 점검 주기: 매일 오전 9시\n\n3. 기록\n점검 결과를 설비 점검일지에 기록한다.\n"""
SAMPLE_NEW = """문서번호: GMP-TR-001\n제목: 정제수 설비 일일 점검 절차 (개정 1)\n\n1. 목적\n정제수 설비의 운전 상태를 매일 확인하고 이상 발생 시 담당자에게 보고한다.\n\n2. 점검 항목\n- 저장탱크 온도: 18~25°C\n- 순환 압력: 1.0~1.5 bar\n- 점검 주기: 매일 오전 8시 30분\n- 경보 표시등 상태를 추가로 확인한다.\n\n3. 기록\n점검 결과를 설비 점검일지에 기록하고, 확인자 서명을 받는다.\n"""


def read_upload(upload) -> str:
    """Extract plain text from supported office document formats."""
    name = upload.name.lower()
    data = upload.getvalue()
    if name.endswith(".txt"):
        return data.decode("utf-8-sig", errors="replace")
    if name.endswith(".docx"):
        from docx import Document
        from docx.table import Table
        from docx.text.paragraph import Paragraph

        doc = Document(BytesIO(data))
        parts = []
        for child in doc.element.body.iterchildren():
            if child.tag.endswith("}p"):
                parts.append(Paragraph(child, doc).text)
            elif child.tag.endswith("}tbl"):
                table = Table(child, doc)
                parts.extend("\t".join(cell.text for cell in row.cells) for row in table.rows)
        return "\n".join(parts)
    if name.endswith(".pdf"):
        from pypdf import PdfReader

        reader = PdfReader(BytesIO(data))
        if reader.is_encrypted:
            raise ValueError("PDF에 암호가 걸려 있습니다. 암호를 해제한 사본을 사용해 주세요.")
        return "\n\n".join(f"[페이지 {i + 1}]\n{page.extract_text() or ''}" for i, page in enumerate(reader.pages))
    if name.endswith(".hwpx"):
        if not zipfile.is_zipfile(BytesIO(data)):
            raise ValueError("올바른 HWPX 파일이 아닙니다.")
        parts = []
        with zipfile.ZipFile(BytesIO(data)) as archive:
            sections = sorted((n for n in archive.namelist() if re.fullmatch(r"Contents/section\d+\.xml", n)), key=lambda n: int(re.search(r"section(\d+)", n).group(1)))
            for section in sections:
                root = ElementTree.fromstring(archive.read(section))
                for paragraph in (node for node in root.iter() if node.tag.split("}")[-1] == "p"):
                    text = "".join(node.text or "" for node in paragraph.iter() if node.tag.split("}")[-1] == "t")
                    parts.append(text)
        return "\n".join(parts)
    if name.endswith(".hwp"):
        # HWP 5.x stores paragraph records in a compressed OLE container.
        import olefile

        ole = olefile.OleFileIO(BytesIO(data))
        try:
            header_data = ole.openstream("FileHeader").read()
            compressed = bool(struct.unpack_from("<I", header_data, 36)[0] & 1)
            streams = sorted((s for s in ole.listdir() if len(s) == 2 and s[0] == "BodyText" and s[1].startswith("Section")), key=lambda s: int(s[1][7:]))
            paragraphs = []
            for stream_name in streams:
                raw = ole.openstream(stream_name).read()
                if compressed:
                    raw = zlib.decompress(raw, -15)
                offset = 0
                while offset + 4 <= len(raw):
                    record = struct.unpack_from("<I", raw, offset)[0]
                    offset += 4
                    tag, size = record & 0x3FF, (record >> 20) & 0xFFF
                    if size == 0xFFF:
                        if offset + 4 > len(raw):
                            break
                        size = struct.unpack_from("<I", raw, offset)[0]
                        offset += 4
                    payload = raw[offset:offset + size]
                    offset += size
                    if tag == 67:  # PARA_TEXT
                        decoded = payload.decode("utf-16le", errors="ignore")
                        # Inline HWP controls may contain binary fields; retain printable text and breaks.
                        decoded = "".join(ch for ch in decoded if ch in "\n\t " or ord(ch) >= 32)
                        paragraphs.append(decoded.strip())
            return "\n".join(paragraphs)
        finally:
            ole.close()
    raise ValueError("지원하지 않는 파일 형식입니다. TXT, DOCX, PDF, HWPX, HWP 파일을 올려 주세요.")


def render_diff(old_text: str, new_text: str) -> tuple[str, int, int]:
    old_lines, new_lines = old_text.splitlines(), new_text.splitlines()
    matcher = difflib.SequenceMatcher(a=old_lines, b=new_lines, autojunk=False)
    old_html, new_html = [], []
    added = sum(j2 - j1 for tag, _, _, j1, j2 in matcher.get_opcodes() if tag in ("insert", "replace"))
    deleted = sum(i2 - i1 for tag, i1, i2, _, _ in matcher.get_opcodes() if tag in ("delete", "replace"))
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            pairs = [(old_lines[i], new_lines[j], "same", "same") for i, j in zip(range(i1, i2), range(j1, j2))]
        elif tag == "delete":
            pairs = [(old_lines[i], "", "deleted", "blank") for i in range(i1, i2)]
        elif tag == "insert":
            pairs = [("", new_lines[j], "blank", "added") for j in range(j1, j2)]
        else:
            pairs = []
            count = max(i2-i1, j2-j1)
            for k in range(count):
                pairs.append((old_lines[i1+k] if i1+k < i2 else "", new_lines[j1+k] if j1+k < j2 else "", "deleted" if i1+k < i2 else "blank", "added" if j1+k < j2 else "blank"))
        for old_line, new_line, old_kind, new_kind in pairs:
            old_prefix = "− " if old_kind == "deleted" else ""
            new_prefix = "+ " if new_kind == "added" else ""
            old_html.append(f'<div class="{old_kind}">{old_prefix}{html.escape(old_line) or "&nbsp;"}</div>')
            new_html.append(f'<div class="{new_kind}">{new_prefix}{html.escape(new_line) or "&nbsp;"}</div>')
    style = "<style>body{font:14px/1.55 sans-serif;margin:0}.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}.col{border:1px solid #ddd;border-radius:8px;overflow:hidden}.head{padding:10px;background:#f2f4f7;font-weight:700}.same,.added,.deleted,.blank{padding:3px 10px;white-space:pre-wrap;overflow-wrap:anywhere;min-height:22px}.added{background:#e5f6e9;color:#165b2b}.deleted{background:#fde8e7;color:#8b1d18}.blank{background:#fafafa}</style>"
    body = '<div class="grid">' + ''.join([f'<div class="col"><div class="head">{name}</div>{"".join(lines)}</div>' for name, lines in (("이전 문서", old_html), ("수정 문서", new_html))]) + '</div>'
    return style + body, added, deleted


def summarize_with_ai(old_text: str, new_text: str) -> str:
    from openai import OpenAI

    client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    model = os.environ.get("OPENAI_MODEL", "gpt-4.1-mini")
    response = client.responses.create(
        model=model,
        instructions="당신은 문서 변경 사항을 요약하는 도우미입니다. 제공된 두 문서의 차이를 한국어로 간결하게 요약하세요. 변경을 추측하거나 규정 준수 여부를 판단하지 말고, 중요한 수치·시간·책임·기록 변경을 구분하세요.",
        input=f"[이전 문서]\n{old_text}\n\n[수정 문서]\n{new_text}",
        max_output_tokens=500,
    )
    return response.output_text


with st.sidebar:
    st.header("문서 입력")
    st.caption("TXT, Word(.docx), PDF, 한글(.hwpx/.hwp) 파일을 올릴 수 있습니다.")
    old_file = st.file_uploader("이전 문서", type=["txt", "docx", "pdf", "hwpx", "hwp"], key="old")
    new_file = st.file_uploader("수정 문서", type=["txt", "docx", "pdf", "hwpx", "hwp"], key="new")
    if st.button("가상 GMP 예시 불러오기", use_container_width=True):
        st.session_state["old_text"] = SAMPLE_OLD
        st.session_state["new_text"] = SAMPLE_NEW
    st.divider()
    ai_enabled = st.checkbox("AI 변경 요약 사용", help="OPENAI_API_KEY 환경 변수가 설정된 경우에만 작동합니다. 사용하면 두 문서의 텍스트가 OpenAI API로 전송됩니다.")

try:
    old_default = read_upload(old_file) if old_file else st.session_state.get("old_text", SAMPLE_OLD)
except Exception as exc:
    old_default = ""
    st.error(f"이전 문서를 읽지 못했습니다: {exc}")
try:
    new_default = read_upload(new_file) if new_file else st.session_state.get("new_text", SAMPLE_NEW)
except Exception as exc:
    new_default = ""
    st.error(f"수정 문서를 읽지 못했습니다: {exc}")
left, right = st.columns(2)
with left:
    old_text = st.text_area("이전 문서 내용", value=old_default, height=260)
with right:
    new_text = st.text_area("수정 문서 내용", value=new_default, height=260)

if st.button("문서 비교하기", type="primary", use_container_width=True):
    if not old_text.strip() and not new_text.strip():
        st.warning("비교할 내용을 입력하거나 파일을 올려 주세요.")
    else:
        diff_html, added, deleted = render_diff(old_text, new_text)
        st.session_state["diff_html"] = diff_html
        st.session_state["counts"] = (added, deleted)
        st.session_state["compared"] = True

if st.session_state.get("compared"):
    added, deleted = st.session_state["counts"]
    a, d = st.columns(2)
    a.metric("추가된 줄", added)
    d.metric("삭제·변경된 줄", deleted)
    st.subheader("변경 내용")
    st.html(st.session_state["diff_html"])
    report = f"# 문서 비교 결과\n\n- 추가된 줄: {added}\n- 삭제·변경된 줄: {deleted}\n\n## 이전 문서\n```text\n{old_text}\n```\n\n## 수정 문서\n```text\n{new_text}\n```\n"
    st.download_button("비교 결과 다운로드 (.md)", report, file_name="document_comparison.md", mime="text/markdown")
    if ai_enabled:
        if not os.environ.get("OPENAI_API_KEY"):
            st.info("AI 요약을 쓰려면 실행 전에 OPENAI_API_KEY 환경 변수를 설정하세요. 일반 비교 기능은 API 키 없이 사용할 수 있습니다.")
        elif st.button("AI로 변경 요약하기"):
            try:
                with st.spinner("요약 중…"):
                    st.session_state["ai_summary"] = summarize_with_ai(old_text, new_text)
            except Exception as exc:
                st.error(f"요약에 실패했습니다. API 키와 네트워크를 확인해 주세요. 상세 내용: {exc}")
        if st.session_state.get("ai_summary"):
            st.subheader("AI 변경 요약")
            st.markdown(st.session_state["ai_summary"])

with st.expander("비교 방식과 한계"):
    st.write("파일에서 글자를 추출한 뒤 줄 단위로 비교합니다. 스캔 이미지 PDF는 글자 추출이 되지 않을 수 있고, PDF의 표나 복잡한 배치는 원래 모양과 다르게 추출될 수 있습니다. 구형 HWP의 표, 글상자, 수식 등은 일부 빠지거나 순서가 달라질 수 있습니다. DOCX 표는 텍스트 형태로 읽습니다. 스캔 PDF는 OCR 기능이 필요합니다.")
