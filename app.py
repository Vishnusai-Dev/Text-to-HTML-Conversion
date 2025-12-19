import streamlit as st
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


# -------------------------------------------------
# Helpers to preserve DOCX structure
# -------------------------------------------------

def iter_blocks(document):
    """Yield paragraphs and tables in exact DOCX order."""
    for child in document.element.body.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, document)
        elif child.tag.endswith('}tbl'):
            yield Table(child, document)


def is_list_paragraph(paragraph):
    pPr = paragraph._p.pPr
    return pPr is not None and pPr.numPr is not None


def is_numbered_list(paragraph):
    numPr = paragraph._p.pPr.numPr
    return numPr is not None and numPr.numId is not None


def render_runs(paragraph):
    """Render paragraph preserving run-level bold."""
    html = ""
    for run in paragraph.runs:
        if run.bold:
            html += f"<strong>{run.text}</strong>"
        else:
            html += run.text
    return html


def is_faq_question(paragraph):
    """
    Detect FAQ questions where:
    - Paragraph is numbered
    - Main text (excluding numbering) is bold
    """
    runs = [run for run in paragraph.runs if run.text.strip()]
    if not runs:
        return False

    # Ignore leading numbering like "1."
    first = runs[0].text.strip()
    if first.rstrip(".").isdigit():
        content_runs = runs[1:]
    else:
        content_runs = runs

    return content_runs and all(run.bold for run in content_runs)


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


# -------------------------------------------------
# Main conversion function
# -------------------------------------------------

def docx_to_html(docx_source):
    document = Document(docx_source)
    html_output = []

    current_list = None  # "ul" or "ol"

    for block in iter_blocks(document):

        # ---------- TABLE ----------
        if isinstance(block, Table):
            if current_list:
                html_output.append(f"</{current_list}>")
                current_list = None
            html_output.append(table_to_html(block))
            continue

        # ---------- PARAGRAPH ----------
        if isinstance(block, Paragraph):
            text = render_runs(block)
            if not text.strip():
                continue

            style = block.style.name if block.style else ""

            # ----- HEADINGS -----
            if style == "Heading 1":
                if current_list:
                    html_output.append(f"</{current_list}>")
                    current_list = None
                html_output.append(f"<h1>{text}</h1>")
                continue

            if style == "Heading 2":
                if current_list:
                    html_output.append(f"</{current_list}>")
                    current_list = None
                html_output.append(f"<h2>{text}</h2>")
                continue

            # ----- LISTS -----
            if is_list_paragraph(block):
                list_type = "ol" if is_numbered_list(block) else "ul"

                if current_list != list_type:
                    if current_list:
                        html_output.append(f"</{current_list}>")
                    html_output.append(f"<{list_type}>")
                    current_list = list_type

                html_output.append(f"<li>{text}</li>")
                continue

            # ----- CLOSE LIST IF NEEDED -----
            if current_list:
                html_output.append(f"</{current_list}>")
                current_list = None

            # ----- FAQ QUESTION -----
            if is_faq_question(block):
                html_output.append(f"<p><strong>{block.text}</strong></p>")
            else:
                html_output.append(f"<p>{text}</p>")

    if current_list:
        html_output.append(f"</{current_list}>")

    return "\n".join(html_output)


# -------------------------------------------------
# Streamlit UI
# -------------------------------------------------

st.set_page_config(page_title="DOCX → HTML Converter", layout="wide")
st.title("DOCX → HTML Converter (Order & Format Exact)")

uploaded_file = st.file_uploader("Upload DOCX file", type=["docx"])

if uploaded_file:
    html_output = docx_to_html(uploaded_file)

    st.subheader("Generated HTML")
    st.code(html_output, language="html")

    st.download_button(
        label="Download HTML",
        data=html_output,
        file_name="output.html",
        mime="text/html"
    )
