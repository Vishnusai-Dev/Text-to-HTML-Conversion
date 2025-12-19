import streamlit as st
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
import re


# -----------------------------------
# Block iterator (order-safe)
# -----------------------------------

def iter_blocks(document):
    for child in document.element.body.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, document)
        elif child.tag.endswith('}tbl'):
            yield Table(child, document)


# -----------------------------------
# Rendering helpers
# -----------------------------------

def render_runs(paragraph):
    html = ""
    for run in paragraph.runs:
        if run.bold:
            html += f"<strong>{run.text}</strong>"
        else:
            html += run.text
    return html


def table_to_html(table):
    html = ["<table border='1' cellpadding='8' cellspacing='0'>"]
    for row in table.rows:
        html.append("<tr>")
        for cell in row.cells:
            cell_html = ""
            for para in cell.paragraphs:
                cell_html += render_runs(para)
            html.append(f"<td>{cell_html}</td>")
        html.append("</tr>")
    html.append("</table>")
    return "\n".join(html)


def is_faq_question_by_text(text):
    """
    FAQ question rule:
    Paragraph starts with 'number + dot'
    Example: '1. Why is ...'
    """
    return bool(re.match(r"^\d+\.\s+", text.strip()))


def is_list_paragraph(paragraph):
    pPr = paragraph._p.pPr
    return pPr is not None and pPr.numPr is not None


# -----------------------------------
# Main conversion
# -----------------------------------

def docx_to_html(docx_source):
    document = Document(docx_source)
    html_output = []

    in_bullet_list = False

    for block in iter_blocks(document):

        # -------- TABLE --------
        if isinstance(block, Table):
            if in_bullet_list:
                html_output.append("</ul>")
                in_bullet_list = False
            html_output.append(table_to_html(block))
            continue

        # -------- PARAGRAPH --------
        if isinstance(block, Paragraph):
            raw_text = block.text.strip()
            if not raw_text:
                continue

            rendered = render_runs(block)
            style = block.style.name if block.style else ""

            # ---- HEADINGS ----
            if style == "Heading 1":
                if in_bullet_list:
                    html_output.append("</ul>")
                    in_bullet_list = False
                html_output.append(f"<h1>{rendered}</h1>")
                continue

            if style == "Heading 2":
                if in_bullet_list:
                    html_output.append("</ul>")
                    in_bullet_list = False
                html_output.append(f"<h2>{rendered}</h2>")
                continue

            # ---- BULLET LIST (ALL LISTS → UL) ----
            if is_list_paragraph(block):
                if not in_bullet_list:
                    html_output.append("<ul>")
                    in_bullet_list = True
                html_output.append(f"<li>{rendered}</li>")
                continue

            # ---- CLOSE LIST IF NEEDED ----
            if in_bullet_list:
                html_output.append("</ul>")
                in_bullet_list = False

            # ---- FAQ QUESTION ----
            if is_faq_question_by_text(raw_text):
                html_output.append(f"<p><strong>{raw_text}</strong></p>")
            else:
                html_output.append(f"<p>{rendered}</p>")

    if in_bullet_list:
        html_output.append("</ul>")

    return "\n".join(html_output)


# -----------------------------------
# Streamlit UI
# -----------------------------------

st.set_page_config(page_title="DOCX → HTML Converter", layout="wide")
st.title("DOCX → HTML Converter (Exact Visual Parity)")

uploaded_file = st.file_uploader("Upload DOCX file", type=["docx"])

if uploaded_file:
    html_output = docx_to_html(uploaded_file)

    st.subheader("Generated HTML")
    st.code(html_output, language="html")

    st.download_button(
        "Download HTML",
        html_output,
        file_name="output.html",
        mime="text/html"
    )

