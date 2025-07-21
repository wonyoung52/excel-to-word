# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.shared import Pt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK
from openai import OpenAI
import math
from io import BytesIO

# ---------------------------- Streamlit 인터페이스 ----------------------------
st.title("엑셀 - 워드 자동 변환기")

api_key = st.text_input("OpenAI API Key", type="password")
uploaded_file = st.file_uploader("Upload (Excel .xlsx)", type=["xlsx"])

# ---------------------------- GPT 처리 함수 정의 ----------------------------
def process_question(number, q_text, client, doc):
    prompt = f"""
문제와 선지를 분리해주세요. 문제는 한 문장, 선지는 '① 선택지내용' 형식으로 출력해주세요. 선지는 최대한 원문 그대로 보존하되 오타나 맞춤법 이상이 있다면 틀린 부분만 교정해주세요.

입력 예시: 다음 중 동물인 것은? 1. 고양이 2. 책상 3. 의자
출력 예시:
문제: 다음 중 동물인 것은?
① 고양이
② 책상
③ 의자

입력: {str(q_text)}
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        content = response.choices[0].message.content.strip()
    except:
        content = f"문제: 오류코드 000"

    lines = content.splitlines()
    question_line = ""
    choices = []

    for l in lines:
        if l.strip().startswith("문제:"):
            question_line = l.replace("문제:", "").strip()
        elif l.strip().startswith(("①", "②", "③", "④", "⑤")):
            choices.append(l.strip())

    try:
        if not question_line:
            raise ValueError("질문 추출 실패")

        para_q = doc.add_paragraph()
        para_q.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        para_q.add_run(f"{number}. {question_line}").bold = True

        if choices:
            para_c = doc.add_paragraph()
            para_c.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            for j, choice in enumerate(choices):
                if j > 0:
                    para_c.add_run().add_break(break_type=WD_BREAK.LINE)
                para_c.add_run(choice).bold = False

        doc.add_paragraph()

    except Exception:
        fail_para = doc.add_paragraph()
        fail_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        fail_para.add_run(f"{number}. 복원 실패").bold = True
        doc.add_paragraph()

# ---------------------------- 문서 생성 조건 ----------------------------
if api_key and uploaded_file:
    client = OpenAI(api_key=api_key)
    df = pd.read_excel(uploaded_file, header=0)
    question_columns = list(range(2, df.shape[1], 3))
    max_q_number = 0

    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Malgun Gothic'
    font.size = Pt(10)
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')

    def add_paragraph(text, bold=False, line_break=False):
        para = doc.add_paragraph()
        run = para.add_run(text)
        run.bold = bold
        para.paragraph_format.line_spacing = Pt(15)
        para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        if line_break:
            run.add_break()
        return para

    for col in question_columns:
        for index in range(len(df)):
            row = df.iloc[index]
            try:
                number_raw = row[col]
                q_text = row[col + 1]

                try:
                    number = int(number_raw)
                    max_q_number = max(max_q_number, number)
                except:
                    number = "?"

                if pd.isna(q_text):
                    add_paragraph(f"{number}. 복원 실패", bold=True)
                    add_paragraph("", line_break=True)
                    continue

                process_question(number, q_text, client, doc)

            except Exception as e:
                add_paragraph(f"복원 실패: {e}", bold=True)
                doc.add_paragraph()

    def add_horizontal_line(paragraph):
        p = paragraph._element
        p_pr = p.get_or_add_pPr()
        border = OxmlElement('w:pBdr')

        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '12')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), 'auto')

        border.append(bottom)
        p_pr.append(border)

    # ------------------------- 정답과 해설 ----------------------------
    add_horizontal_line(doc.add_paragraph())
    doc.add_paragraph().add_run("정답 및 해설").bold = True
    total_rows = math.ceil(max_q_number / 5 * 2) + 1
    table = doc.add_table(rows=total_rows, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tbl_borders.append(border)
    tbl_pr.append(tbl_borders)

    for i in range(total_rows):
        row = table.rows[i]
        is_shaded = i % 2 == 0
        for j in range(5):
            cell = row.cells[j]
            para = cell.paragraphs[0]
            para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = para.add_run()
            if is_shaded:
                value = i // 2 * 5 + j + 1
                if value <= max_q_number:
                    run.text = str(value)
                shading_elm = OxmlElement("w:shd")
                shading_elm.set(qn("w:val"), "clear")
                shading_elm.set(qn("w:fill"), "D9D9D9")
                cell._tc.get_or_add_tcPr().append(shading_elm)

    # ------------------------- 마무리 문단 ----------------------------
    add_horizontal_line(doc.add_paragraph())
    add_paragraph("강의 총평", bold=True)
    add_horizontal_line(doc.add_paragraph())
    add_paragraph("족관 총평", bold=True)
    add_horizontal_line(doc.add_paragraph())
    add_paragraph("시험 난이도", bold=True)

    # ------------------------- 다운로드 ----------------------------

    if "doc_buffer" not in st.session_state:
        with st.spinner("Generating..."):
            buffer = BytesIO()
            doc.save(buffer)
            st.session_state["doc_buffer"] = buffer
    else:
        buffer = st.session_state["doc_buffer"]

    st.download_button(
        label="Download",
        data=buffer.getvalue(),
        file_name="족보.docx",
     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )