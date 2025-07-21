# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK
from openai import OpenAI
import math
from io import BytesIO

# ---------------------------- Streamlit UI ----------------------------
st.title("엑셀 - 워드 자동 변환기")

with st.expander("사용법 펼치기"):
    st.markdown("""                
    1. OpenAI API Key를 입력하세요.

    2. 족보 제목을 입력하세요.
    - 문서 맨 위에 삽입될 제목입니다.
    - 기본값: 2024 족보

    3. 저장할 파일명을 입력하세요.       
    - 저장되는 파일명입니다. 확장자는 자동으로 입력됩니다.
    - 기본값: 제목 없음
                
    4.  엑셀 파일 업로드 (.xlsx)
    - 엑셀 파일을 업로드하세요.
    - 반드시 문제 번호에 대하여 오름차순 정렬을 시행한 후 업로드 해주세요.

    5. 다운로드
    - 파일을 업로드 하면 변환이 자동으로 실행됩니다.
    - 실행 시간은 문제당 1-2초입니다.

    6. 유의사항
    - 문제 번호가 정수가 아니면 "?"로 표시됩니다.
    - 문제 내용 항목이 비어있는 경우 "복원 실패"로 표시됩니다.  
    - 문제 내용 항목은 GPT-4o를 통해 자동 분석되므로 API 사용량에 유의하세요.
    """)

api_key = st.text_input("OpenAI API Key", type="password")
uploaded_file = st.file_uploader("Upload (Excel .xlsx)", type=["xlsx"])
title_input = st.text_input("Title", value="2024 족보")
file_name_input = st.text_input("File Name", value="제목 없음")

# ---------------------------- GPT 처리 함수 ----------------------------
def add_paragraph(text, bold=False, line_break=False):
    para = doc.add_paragraph()
    run = para.add_run(text)
    run.bold = bold
    para.paragraph_format.line_spacing = Pt(15)
    para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    if line_break:
        run.add_break()
    return para

def process_question(number, q_text, client, doc):
    prompt = f"""
문제와 선지를 분리해주세요. 문제는 한 문장, 선지는 '① 선택지내용' 형식으로 출력해주세요. 선지는 최대한 원문 그대로 보존하되 오타나 맞춤법 이상이 있다면 틀린 부분만 교정해주세요.

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
                para_c.add_run(choice)

        doc.add_paragraph()

    except Exception:
        fail_para = doc.add_paragraph()
        fail_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        fail_para.add_run(f"{number}. 복원 실패").bold = True
        doc.add_paragraph()

# ---------------------------- 실행 조건 ----------------------------
if "converted" not in st.session_state:
    st.session_state["converted"] = False

if st.button("Convert") and not st.session_state.get("converted", False):
    st.session_state["converted"] = True

    with st.spinner("Generating..."):
        client = OpenAI(api_key=api_key)
        df = pd.read_excel(uploaded_file, header=0)

        if "사진 자료" in df.columns:
            photo_col_idx = df.columns.get_loc("사진 자료")
        else:
            photo_col_idx = df.shape[1]
        
        question_columns = [i for i in range(2, photo_col_idx, 3)]
        max_q_number = 0

        doc = Document()
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Malgun Gothic'
        font.size = Pt(10)
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')

        # ---------------------------- 제목 ----------------------------
        title_para = doc.add_paragraph()
        title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        title_run = title_para.add_run(title_input)
        title_run.font.size = Pt(16)
        title_run.font.name = 'Malgun Gothic'
        title_run.bold = True
        title_run.underline = True
        doc.add_paragraph()

        # ---------------------------- 교수정보 표 ----------------------------
        table = doc.add_table(rows=5, cols=3)
        table.style = 'Table Grid'
        headers = ["교수님", "번호", "담당 강의 (시수)"]
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            para = cell.paragraphs[0]
            para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = para.add_run(header)
            run.bold = True
            run.font.size = Pt(10)
        doc.add_paragraph()

        # ---------------------------- 문제 처리 ----------------------------
        for col in question_columns:
            for index in range(len(df)):
                row = df.iloc[index]
                try:
                    number_raw = row[col]
                    q_text = row[col + 1]
                    number = int(number_raw) if not pd.isna(number_raw) else "?"
                    max_q_number = max(max_q_number, number if isinstance(number, int) else 0)

                    if pd.isna(q_text):
                        doc.add_paragraph(f"{number}. 복원 실패").bold = True
                        doc.add_paragraph()
                        continue

                    process_question(number, q_text, client, doc)

                except Exception as e:
                    doc.add_paragraph(f"복원 실패: {e}")
                    doc.add_paragraph()

        # ---------------------------- 해설 표 ----------------------------
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

        add_horizontal_line(doc.add_paragraph())
        doc.add_paragraph().add_run("정답 및 해설").bold = True
        total_rows = math.ceil(max_q_number / 5 * 2)
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

            tr = row._tr
            tr_pr = tr.get_or_add_trPr()
            tr_height = OxmlElement('w:trHeight')
            tr_height.set(qn('w:val'), '400')        
            tr_height.set(qn('w:hRule'), 'exact')
            tr_pr.append(tr_height)

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

        # ---------------------------- 총평 ----------------------------
        add_horizontal_line(doc.add_paragraph())
        add_paragraph("강의 총평", bold=True)
        add_horizontal_line(doc.add_paragraph())
        add_paragraph("족관 총평", bold=True)
        add_horizontal_line(doc.add_paragraph())
        add_paragraph("시험 난이도", bold=True)

        # ---------------------------- 다운로드 ----------------------------
        buffer = BytesIO()
        doc.save(buffer)

        st.download_button(
            label="Download Word File",
            data=buffer.getvalue(),
            file_name=f"{file_name_input}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )