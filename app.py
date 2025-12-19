import streamlit as st
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_blocks(document):
    for child in document.element.body.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, document)
        elif child.tag.endswith('}tbl'):
            yield Table(child, document)


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


def docx_to_html(docx_source):
    document = Document(docx_source)
    html_output = []

    for block in iter_blocks(document):
        if isinstance(block, Paragraph):
            content = render_runs(block)
            if not content.strip():
                continue

            style_name = block.style.name if block.style else ""

            if style_name == "Heading 1":
                html_output.append(f"<h1>{content}</h1>")
            elif style_name == "Heading 2":
                html_output.append(f"<h2>{content}</h2>")
            else:
                html_output.append(f"<p>{content}</p>")

        elif isinstance(block, Table):
            html_output.append(table_to_html(block))

    return "\n".join(html_output)


# ---------------- Streamlit UI ----------------

st.set_page_config(page_title="DOCX → HTML Converter", layout="wide")
st.title("DOCX → HTML Converter (Order, Format & Headers Exact)")

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


