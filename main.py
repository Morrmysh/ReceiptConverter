import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import datetime

def excel_to_pdf(excel_file_path):
    try:
        wb = load_workbook(excel_file_path)
        sheet = wb.active

        current_date = datetime.datetime.now().strftime('%Y-%m-%d')

        output_folder = os.path.join(os.path.dirname(excel_file_path), "Квитанции от " + current_date)
        os.makedirs(output_folder, exist_ok=True)
        pdf_file_path = os.path.join(str(output_folder), str(os.path.basename(excel_file_path).replace('.xlsx', '.pdf')))

        pdf = SimpleDocTemplate(str(pdf_file_path), pagesize=landscape(A4))
        pdf.topMargin = 10
        pdf.bottomMargin = 10
        pdf.leftMargin = 20
        pdf.rightMargin = 20

        pdfmetrics.registerFont(TTFont('TimesNewRoman', 'assets/times.ttf'))
        pdfmetrics.registerFont(TTFont('TimesNewRoman-Bold', 'assets/timesbd.ttf'))
        header_style = ParagraphStyle(
            name='HeaderStyle',
            fontName='TimesNewRoman',
            fontSize=14,
            spaceAfter=12,
        )
        text_style = ParagraphStyle(
            name='TextStyle',
            fontName='TimesNewRoman',
            fontSize=12,
            spaceAfter=10,
            spaceBefore=10,
        )
        cell_style = ParagraphStyle(
            name='TextStyle',
            fontName='TimesNewRoman',
            fontSize=10,
            spaceAfter=10,
            wordWrap=None,  # 'CJK'
            alignment=0,
        )
        cell_style_bold = ParagraphStyle(
            name='TextStyle',
            fontName='TimesNewRoman-Bold',
            fontSize=10,
            spaceAfter=10,
            wordWrap=None,  # 'CJK'
            alignment=0,
        )
        center_cell_style = ParagraphStyle(
            name='TextStyle',
            fontName='TimesNewRoman',
            fontSize=10,
            spaceAfter=10,
            wordWrap=None,  # 'CJK'
            alignment=1,
        )
        center_cell_style_bold = ParagraphStyle(
            name='TextStyle',
            fontName='TimesNewRoman-Bold',
            fontSize=10,
            spaceAfter=10,
            wordWrap=None,  # 'CJK'
            alignment=1,
        )

        content = []

        title = Paragraph(
            f"Квитанция номер {sheet['B4'].value.strip()}.",
            header_style,
        )
        content.append(title)

        excel_data = [[str(cell.value) for cell in row] for row in sheet['A5:D8']]
        excel_data.insert(1, ["ИНН:", "5042043326", "", ""])
        data = []
        for row in excel_data:
            data.append([])
            for num, cell in enumerate(row):
                if num % 2 == 0:
                    data[-1].append(Paragraph(cell, cell_style_bold))
                else:
                    data[-1].append(Paragraph(cell, cell_style))

        info_table = Table(data, style=[
            ('SPAN', (2, 0), (2, 1)),
            ('SPAN', (3, 0), (3, 1)),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('LINEABOVE', (0, 0), (-1, -1), 1, colors.black),
            ('LINEAFTER', (1, 0), (1, -1), 1, colors.black)
        ],
                           colWidths=[120, None, 140, None],
                           hAlign='LEFT')
        content.append(info_table)

        text = Paragraph("Расшифровка по видам услуг", text_style)
        content.append(text)

        excel_data = [[str(cell.value).replace('_x000D_\n', '').replace('None', '') for cell in row] for row in sheet['A12:P25']]
        data = []
        for row in excel_data[:3]:
            data.append([])
            for cell in row:
                data[-1].append(Paragraph(cell, center_cell_style_bold))
        for row in excel_data[3:]:
            data.append([])
            data[-1].append(Paragraph(row[0], cell_style_bold))
            for cell in row[1:]:
                data[-1].append(Paragraph(cell, center_cell_style))

        summary_table = Table(data, style=[
            ('SPAN', (0, 0), (0, 2)),
            ('SPAN', (1, 0), (1, 2)),
            ('SPAN', (2, 0), (2, 2)),
            ('SPAN', (3, 0), (3, 2)),
            ('SPAN', (4, 0), (4, 2)),
            ('SPAN', (5, 0), (8, 0)),
            ('SPAN', (5, 1), (6, 1)),
            ('SPAN', (7, 1), (8, 1)),
            ('SPAN', (5, 2), (7, 2)),
            ('SPAN', (9, 0), (9, 2)),
            ('SPAN', (10, 0), (11, 0)),
            ('SPAN', (10, 1), (10, 2)),
            ('SPAN', (11, 1), (11, 2)),
            ('SPAN', (12, 0), (13, 0)),
            ('SPAN', (12, 1), (12, 2)),
            ('SPAN', (13, 1), (13, 2)),
            ('SPAN', (14, 0), (14, 2)),
            ('SPAN', (15, 0), (15, 2)),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BACKGROUND', (0, 0), (-1, 2), colors.lightgrey),
            ('BACKGROUND', (0, 3), (-1, 4), colors.whitesmoke),
            ('BACKGROUND', (0, 6), (-1, 6), colors.whitesmoke),
            ('BACKGROUND', (0, 12), (-1, 12), colors.whitesmoke),
            ('BACKGROUND', (0, 3), (0, -1), colors.whitesmoke),
        ],
                              colWidths=[100, 60, 60, 50, 40, 30, 55, 30, 40, 40, 40, 40, 60, None, None, None],)
        content.append(summary_table)

        data = [[Paragraph('Мясникова Ангелина Михайловна', text_style),
                Paragraph(datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), text_style)]]
        table = Table(data, style=[('TOPPADDING', (0, 0), (-1, -1), 10), ('LEFTPADDING', (0, 0), (-1, -1), 0)])
        content.append(table)

        pdf.build(content)

        return pdf_file_path

    except Exception as e:
        tk.messagebox.showerror("Error", str(e))
    # finally:
    #     pass


def open_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])
    if file_path:
        pdf_file_path = excel_to_pdf(file_path)
        if pdf_file_path:
            tk.messagebox.showinfo("Success", f"Квитанция создана: {pdf_file_path}")


def open_overhaul():
    pass


# Создание GUI
root = tk.Tk()
root.title("Receipt Converter")
root.geometry("300x150")

# Кнопка для выбора файла
overhaul_button = tk.Button(root, text="Капремонт", command=open_overhaul)
overhaul_button.pack(pady=20)
receipt_button = tk.Button(root, text="Коммунальные платежи", command=open_file)
receipt_button.pack(pady=20)

# Запуск GUI
root.mainloop()