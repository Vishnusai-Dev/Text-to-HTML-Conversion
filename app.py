import streamlit as st
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


# -------------------------------
# Core DOCX → HTML logic
# -------------------------------

def iter_blocks(document):
    """
    Yield paragraphs and tables in the exact order
    they appear in the DOCX file.
    """
    for child in document.element.body.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, document)
        elif child.tag.endswith('}tbl'):
            yield Table(child, document)


def render_runs(paragraph):
    """
    Render paragraph text preserving run-level bold.
    No inferred formatting.
    """
    html = ""
    for run in paragraph.runs:
        if run.bold:
            html += f"<strong>{run.text}</strong>"
        else:
            html += run.text
    return html


def table_to_html(table):
    """
    Convert a Word table to HTML with exact inline formatting.
    """
    html = []
    html.append("<table border='1' cellpadding='8' cellspacing='0'>")

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


def docx_to_html(docx_source):
    """
    Convert DOCX to HTML.
    docx_source can be:
    - Streamlit UploadedFile
    - file-like object
    - file path (string)
    """
    document = Document(docx_source)
    html_output = []

    for block in iter_blocks(document):
        if isinstance(block, Paragraph):
            content = render_runs(block)
            if content.strip():
                html_output.append(f"<p>{content}</p>")

        elif isinstance(block, Table):
            html_output.append(table_to_html(block))

    return "\n".join(html_output)


# -------------------------------
# Streamlit UI
# -------------------------------

st.set_page_config(page_title="DOCX → HTML Converter", layout="wide")
st.title("DOCX → HTML Converter (Order & Format Exact)")

uploaded_file = st.file_uploader(
    "Upload a DOCX file",
    type=["docx"]
)

if uploaded_file:
    try:
        html_output = docx_to_html(uploaded_file)

        st.subheader("Generated HTML")
        st.code(html_output, language="html")

        st.download_button(
            label="Download HTML",
            data=html_output,
            file_name="output.html",
            mime="text/html"
        )

    except Exception as e:
        st.error("Failed to process the document.")
        st.exception(e)

