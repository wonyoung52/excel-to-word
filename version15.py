# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from docx.oxml.ns import qn
from docx.shared import Pt
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK
from openai import OpenAI
from io import BytesIO

st.title("엑셀 - 워드 자동 변환기")

api_key = st.text_input("OpenAI API Key", type="password")
uploaded_file = st.file_uploader("Upload (Excel .xlsx)", type=["xlsx"])
title_input = st.text_input("Title", value="2024 족보")
file_name_input = st.text_input("File Name", value="제목 없음")


# 🔥 GPT 응답 안전 파싱 함수
def get_text_from_response(response):
    try:
        return response.output[0].content[0].text
    except:
        try:
            return response.output_text
        except:
            return ""


# ---------------- GPT 처리 ----------------
def process_question(number, q_text, client, doc):

    prompt = f"""
문제와 선지를 분리해주세요.
반드시 아래 형식으로 출력:
문제: ...
① ...
② ...
③ ...
④ ...
⑤ ...

입력:
{str(q_text)}
"""

    try:
        response = client.responses.create(
            model="gpt-5.3-chat-latest",   # 🔥 가장 안정적
            input=prompt
        )

        content = get_text_from_response(response).strip()

        # 🔥 디버깅 (문제 있으면 이거 보고 바로 알 수 있음)
        # st.write(content)

    except Exception as e:
        st.error(e)
        content = "문제: 오류"

    lines = content.splitlines()

    question_line = ""
    choices = []

    for l in lines:
        s = l.strip()
        if s.startswith("문제:"):
            question_line = s.replace("문제:", "").strip()
        elif s.startswith(("①", "②", "③", "④", "⑤")):
            choices.append(s)

    # 🔥 fallback (핵심)
    if not question_line:
        question_line = lines[0] if lines else "복원 실패"

    # ---------------- docx 작성 ----------------
    try:
        para_q = doc.add_paragraph()
        para_q.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        para_q.add_run(f"{number}. {question_line}").bold = True

        if choices:
            para_c = doc.add_paragraph()
            para_c.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            for j, choice in enumerate(choices):
                if j > 0:
                    para_c.add_run().add_break(WD_BREAK.LINE)
                para_c.add_run(choice)

        doc.add_paragraph()

    except Exception as e:
        p = doc.add_paragraph()
        p.add_run(f"{number}. 복원 실패 ({e})").bold = True
        doc.add_paragraph()


# ---------------- 실행 ----------------
if st.button("Convert"):

    if not api_key:
        st.error("API Key 입력하세요")
        st.stop()

    if uploaded_file is None:
        st.error("엑셀 파일 업로드하세요")
        st.stop()

    with st.spinner("Generating..."):

        client = OpenAI(api_key=api_key)
        df = pd.read_excel(uploaded_file)

        doc = Document()

        style = doc.styles['Normal']
        font = style.font
        font.name = 'Malgun Gothic'
        font.size = Pt(10)
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')

        # 제목
        title_para = doc.add_paragraph()
        title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = title_para.add_run(title_input)
        run.bold = True
        run.font.size = Pt(16)
        doc.add_paragraph()

        # ---------------- 문제 처리 ----------------
        max_q_number = 0

        for i in range(len(df)):
            try:
                number = int(df.iloc[i, 0]) if not pd.isna(df.iloc[i, 0]) else "?"
                q_text = df.iloc[i, 1]

                if pd.isna(q_text):
                    p = doc.add_paragraph()
                    p.add_run(f"{number}. 복원 실패").bold = True
                    doc.add_paragraph()
                    continue

                process_question(number, q_text, client, doc)

            except Exception as e:
                p = doc.add_paragraph()
                p.add_run(f"복원 실패: {e}").bold = True
                doc.add_paragraph()

        # ---------------- 다운로드 ----------------
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="Download",
            data=buffer.getvalue(),
            file_name=f"{file_name_input}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )