import io
import os
from flask import Flask, render_template, request, send_file, redirect, flash, url_for
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

app = Flask(__name__)
app.secret_key = 'change-me'  # needed for flash messages
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # max upload: 16 MB

MAX_ROWS = 5000  # safety limit so giant files don't freeze your browser


def read_csv_safely(file_storage, delimiter):
    """Read the uploaded CSV into a pandas DataFrame with sensible fallbacks."""
    # rewind just in case
    file_storage.stream.seek(0)
    # Auto-detect delimiter if 'auto'; else use the one user chose
    sep = None if delimiter == 'auto' else (delimiter if delimiter else None)

    try:
        # engine='python' allows auto detection when sep=None
        df = pd.read_csv(file_storage.stream, sep=sep, nrows=MAX_ROWS, engine='python')
    except UnicodeDecodeError:
        # try a different encoding
        file_storage.stream.seek(0)
        df = pd.read_csv(file_storage.stream, sep=sep, nrows=MAX_ROWS, engine='python', encoding='latin-1')
    except Exception:
        # last resort: try some common encodings
        df = None
        for enc in ['utf-8', 'utf-16', 'cp1252', 'latin-1']:
            try:
                file_storage.stream.seek(0)
                df = pd.read_csv(file_storage.stream, sep=sep, nrows=MAX_ROWS, engine='python', encoding=enc)
                break
            except Exception:
                df = None
        if df is None:
            raise  # bubble up the error
    return df


def df_to_pdf_bytes(df, title='CSV to PDF', landscape_mode=True):
    """Turn a DataFrame into a neat PDF table and return as BytesIO."""
    buffer = io.BytesIO()
    page_size = landscape(A4) if landscape_mode else A4

    # page/margins
    left, right, top, bottom = 20, 20, 30, 20
    doc = SimpleDocTemplate(buffer, pagesize=page_size,
                            leftMargin=left, rightMargin=right,
                            topMargin=top, bottomMargin=bottom)

    elements = []
    styles = getSampleStyleSheet()
    title_style = styles['Heading1']
    title_style.alignment = 1  # center
    elements.append(Paragraph(title, title_style))
    elements.append(Spacer(1, 10))

    # cell styles
    cell_style = ParagraphStyle('cell', fontSize=8, leading=10)
    header_style = ParagraphStyle('header', fontSize=9, leading=11)

    # Turn NaN into blanks
    df = df.fillna('')

    # headers
    headers = [Paragraph(str(col), header_style) for col in df.columns.tolist()]

    # rows
    data_rows = []
    for _, row in df.iterrows():
        cells = []
        for val in row.tolist():
            text = str(val)
            # escape special characters for XML/HTML
            text = (text.replace('&', '&amp;')
                        .replace('<', '&lt;')
                        .replace('>', '&gt;')
                        .replace('\n', '<br/>'))
            cells.append(Paragraph(text, cell_style))
        data_rows.append(cells)

    table_data = [headers] + data_rows

    # column widths: proportional to header length (simple heuristic)
    usable_width = page_size[0] - left - right
    num_cols = max(1, len(headers))
    header_lengths = [max(5, len(h.getPlainText())) for h in headers]
    total_len = sum(header_lengths) if sum(header_lengths) else num_cols
    col_widths = [max(40, usable_width * (l / total_len)) for l in header_lengths]

    table = Table(table_data, colWidths=col_widths, repeatRows=1)

    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f0f0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#aaaaaa')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#fbfbfb')]),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
    ]))

    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    file = request.files.get('csv_file')
    delimiter = request.form.get('delimiter', 'auto')
    title = request.form.get('title', 'CSV to PDF')
    orientation = request.form.get('orientation', 'landscape')

    if not file or file.filename == '':
        flash('Please choose a CSV file.')
        return redirect(url_for('index'))

    if not file.filename.lower().endswith('.csv'):
        flash('Only .csv files are allowed.')
        return redirect(url_for('index'))

    try:
        df = read_csv_safely(file, delimiter)
        if df.shape[0] == 0:
            flash('Your CSV appears to be empty.')
            return redirect(url_for('index'))

        pdf_bytes = df_to_pdf_bytes(df, title=title.strip() or 'CSV to PDF',
                                    landscape_mode=(orientation == 'landscape'))

        filename = os.path.splitext(file.filename)[0] + '.pdf'
        return send_file(pdf_bytes, as_attachment=True,
                         download_name=filename, mimetype='application/pdf')

    except Exception as e:
        print('Error during conversion:', e)
        flash('Sorry, something went wrong while converting your file. Make sure it is a valid CSV.')
        return redirect(url_for('index'))


if __name__ == '__main__':
    # Tip: change host='0.0.0.0' to test on phone in same Wi-Fi
    app.run(debug=True)
