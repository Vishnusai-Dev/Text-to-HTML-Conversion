
import streamlit as st
from docx import Document
import difflib
from io import StringIO

st.set_page_config(page_title="DOCX to HTML Converter", layout="wide")

st.title("DOCX → HTML Converter (Format-Exact)")

uploaded_file = st.file_uploader("Upload DOCX file", type=["docx"])

def docx_to_html(doc):
    html = []
    for para in doc.paragraphs:
        line = ""
        for run in para.runs:
            if run.bold:
                line += f"<strong>{run.text}</strong>"
            else:
                line += run.text
        if line.strip():
            html.append(f"<p>{line}</p>")

    for table in doc.tables:
        html.append("<table border='1' cellpadding='8' cellspacing='0'>")
        for row in table.rows:
            html.append("<tr>")
            for cell in row.cells:
                cell_text = ""
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.bold:
                            cell_text += f"<strong>{run.text}</strong>"
                        else:
                            cell_text += run.text
                html.append(f"<td>{cell_text}</td>")
            html.append("</tr>")
        html.append("</table>")
    return "\n".join(html)

if uploaded_file:
    doc = Document(uploaded_file)
    html_output = docx_to_html(doc)

    st.subheader("Generated HTML")
    st.code(html_output, language="html")

    st.download_button(
        "Download HTML",
        html_output,
        file_name="output.html",
        mime="text/html"
    )

    st.subheader("DOC vs HTML Diff Checker")
    doc_text = "\n".join([p.text for p in doc.paragraphs])

    diff = difflib.HtmlDiff().make_file(
        doc_text.splitlines(),
        html_output.splitlines(),
        fromdesc="DOCX Text",
        todesc="HTML Output"
    )

    st.components.v1.html(diff, height=600, scrolling=True)
