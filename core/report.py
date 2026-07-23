import io
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

def generate_pdf_from_view_df(order_df, today_str, provider_name="지오영"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=30, bottomMargin=30)
    story = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontSize=16, spaceAfter=15, alignment=1
    )

    story.append(Paragraph(f"[{today_str}] {provider_name} 입고 검수 현황 보고서", title_style))
    story.append(Spacer(1, 10))

    table_data = [["No", "제품코드", "제품전산출력명", "규격", "수량", "스캔", "상태"]]
    total_expected = total_scanned = 0

    for _, row in order_df.iterrows():
        expected = int(row["수량"])
        scanned = int(row["스캔수량"])
        is_registered = row["등록여부"]
        total_expected += expected
        total_scanned += scanned
        
        if not is_registered:
            status = "미등록"
        elif scanned == expected and expected > 0:
            status = "완료"
        elif scanned < expected:
            status = f"미달({scanned-expected})"
        else:
            status = f"초과(+{scanned-expected})"

        table_data.append([
            str(row["No"]), str(row["제품코드"]), str(row["제품전산출력명"])[:16],
            str(row["제품규격"]), str(expected), str(scanned), status
        ])

    table_data.append(["총계", "-", "-", "-", str(total_expected), str(total_scanned), "-"])

    pdf_table = Table(table_data, colWidths=[30, 60, 180, 70, 45, 45, 65])
    pdf_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#EAEDED")),
    ]))

    story.append(pdf_table)
    doc.build(story)
    buffer.seek(0)
    return buffer