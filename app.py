import streamlit as st
import re
from io import BytesIO
from docx import Document
from docx.shared import Inches

st.set_page_config(page_title="MCQ Formatter", layout="wide")

st.title("📚 Bhargav's MCQ Formatter (English & Assamese)")

raw_text = st.text_area(
    "Paste Questions Here (Supports Single-line & Multi-line Options)",
    height=400
)

if st.button("Format MCQs"):

    # Smart Splitter: Recognizes prefixes (Q1., প্ৰশ্ন ১.) OR standalone multi-digit question numbers (121.)
    # This safely ignores single-digit statement lists (1., 2.) inside a question body.
    blocks = re.split(r'\n(?=(?:Q|প্ৰশ্ন|প্ৰ\.)\s*(?:\d+|[০-৯]+)[\.\)]|\s*(?:\d{2,}|[০-৯]{2,})[\.\)])', raw_text.strip())

    doc = Document()

    # Answer translation map
    answer_map = {
        "A": "1", "B": "2", "C": "3", "D": "4",
        "ক": "1", "খ": "2", "গ": "3", "ঘ": "4"
    }

    total_questions = 0

    for block in blocks:
        if not block.strip():
            continue

        # Extract textual context using holistic block mapping instead of line-by-line loops
        block_text = block.strip()

        # 1. Identify and extract the Answer and any trailing Solution text
        ans_match = re.search(r'(?:correct\s+)?(answer|উত্তৰ)\s*:\s*\(?([a-dA-Dকখগঘ])[\.\)]?', block_text, re.IGNORECASE)
        
        answer = ""
        solution = ""
        block_minus_answer = block_text

        if ans_match:
            answer = ans_match.group(2).upper()
            ans_start = ans_match.start()
            solution = block_text[ans_match.end():].strip()
            # Isolate the question and options away from the answer tag
            block_minus_answer = block_text[:ans_start].strip()

        # 2. Extract choice blocks based on regex boundary markers
        opt_matches = list(re.finditer(r'(?:^|\s|\b)(\(?[A-Da-dকখগঘ][\.\)])(?=\s|$)', block_minus_answer))
        
        question = ""
        options = []

        if opt_matches:
            # The question text is everything structural leading up to the very first choice item
            question = block_minus_answer[:opt_matches[0].start()].strip()
            
            # Slice segments dynamically between choice keys
            for i in range(len(opt_matches)):
                start_pos = opt_matches[i].end()
                end_pos = opt_matches[i+1].start() if i + 1 < len(opt_matches) else len(block_minus_answer)
                opt_text = block_minus_answer[start_pos:end_pos].strip()
                if opt_text:
                    options.append(opt_text)
        else:
            # Fallback if no choice elements are targeted
            question = block_minus_answer

        # 3. Apply Clean Up Rules to the Question Field
        # Strip leading numbers/headers (e.g., "121.", "Q99.")
        question = re.sub(r'^(?:Q|প্ৰশ্ন|প্ৰ\.)?\s*(?:\d+|[০-৯]+)[\.\s)]*', '', question)
        # Strip leading parenthetical metadata blocks
        question = re.sub(r'^\([^)]*\)\s*', '', question)
        # Force individual statement tags (I., II., 1., ২.) down onto distinct lines
        question = re.sub(r'\s+(?=(?:[IVXivx]+|\d+|[০-৯]+)[\.\)]\s|\((?:[IVXivx]+|\d+|[০-৯]+)\)\s)', '\n', question)
        # Force prompt inquiries ("Which of...", "তলৰ...") onto their own final line
        question = re.sub(r'\s+(?=(?:Which of|Choose the|Select the|ওপৰৰ|তলৰ|কোনটো|কোনবোৰ)(?:\s|$))', '\n', question)

        # Convert option characters to output numbers
        answer_numeric = answer_map.get(answer, answer)

        # Build Output Table Architecture
        table = doc.add_table(rows=0, cols=2)
        table.style = "Table Grid"
        table.columns[0].width = Inches(1.3)
        table.columns[1].width = Inches(5.8)

        def add_row(label, value):
            row = table.add_row().cells
            row[0].text = label
            row[1].text = value
            row[0].width = Inches(1.3)
            row[1].width = Inches(5.8)

        # Assign records into target fields
        add_row("Question", question)
        add_row("Type", "multiple_choice")

        for option in options:
            add_row("Option", option)

        add_row("Answer", answer_numeric)
        add_row("Solution", solution)
        add_row("Positive Marks", "1")
        add_row("Negative Marks", "0.25")

        doc.add_paragraph()
        total_questions += 1

    buffer = BytesIO()
    doc.save(buffer)

    st.success(f"{total_questions} questions detected.")

    st.download_button(
        "📥 Download Word",
        data=buffer.getvalue(),
        file_name="formatted_mcqs.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
