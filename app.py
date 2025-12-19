from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_blocks(document):
    """
    Yield paragraphs and tables in the exact order
    they appear in the DOCX.
    """
    body = document.element.body
    for child in body.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, document)
        elif child.tag.endswith('}tbl'):
            yield Table(child, document)


def render_runs(paragraph):
    """
    Render paragraph text preserving run-level bold.
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


def docx_to_html(docx_path, output_html_path):
    """
    Main conversion function.
    """
    doc = Document(docx_path)
    html_output = []

    for block in iter_blocks(doc):
        if isinstance(block, Paragraph):
            content = render_runs(block)
            if content.strip():
                html_output.append(f"<p>{content}</p>")

        elif isinstance(block, Table):
            html_output.append(table_to_html(block))

    html = "\n".join(html_output)

    with open(output_html_path, "w", encoding="utf-8") as f:
        f.write(html)


# -------------------------------
# Example usage
# -------------------------------
if __name__ == "__main__":
    input_docx = "input.docx"
    output_html = "output.html"
    docx_to_html(input_docx, output_html)
