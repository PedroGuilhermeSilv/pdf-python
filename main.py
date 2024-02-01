from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import io
import textwrap
import pandas as pd

def create_canvas(textos):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.setFont("Helvetica", 11)
    for key, value in textos.items():
        if pd.notna(value):
            textos[key] = str(value)
        else:
            textos[key] = ''
    return can, packet, textos

def write_pdf(existing_pdf, packet, output_path):
    packet.seek(0)
    new_pdf = PdfReader(packet)
    for i, page in enumerate(existing_pdf.pages):
        if i < len(new_pdf.pages):
            page.merge_page(new_pdf.pages[i])
    output = PdfWriter()
    for page in existing_pdf.pages:
        output.add_page(page)
    with open(output_path, "wb") as outputStream:
        output.write(outputStream)

def fill_pdf_template(textos, output_path):
    existing_pdf = PdfReader(open("molde_descritivo_cargo.pdf", "rb"))
    can, packet, textos = create_canvas(textos)

    # Página 1
    can.drawString(150, 665, textos['nome'])
    can.drawString(280, 630, textos['setor'])
    can.drawString(37, 614, textos['cargo'])
    can.drawString(43, 574, textos['superior_imediato'])
    can.setFont("Helvetica", 8)
    can.drawString(120, 552, textos['local_trabalho'])
    can.setFont("Helvetica", 11)
    can.drawString(362, 550, textos['ch'])
    lines = textwrap.wrap(textos['objetivo_cargo'], width=98)
    match textos['regime_contratacao']:
        case 'Comissionado':
            can.drawString(324.6, 584, '\n')
        case 'Seletivo':
            can.drawString(424, 584, '\n')
        case 'Efetivo':
            can.drawString(494, 584, '\n')
    y = 519
    for line in lines:
        can.drawString(34, y, line)
        y -= 13
    y = 320
    for i in range(1, 8):
        if textos[f'descricao_sumaria_0{i}']:
            can.drawString(35, y, " • "+textos[f'descricao_sumaria_0{i}'])
        y -= 15
    can.showPage()  # Finaliza a página 1 e começa uma nova página

    # Página 2
    y=640
    can.setFont("Helvetica", 11)
    for i in range (1,6):
        if textos[f'descricao_compet_comportamentais_0{i}']:
            can.drawString(35, y, " • "+textos[f'descricao_compet_comportamentais_0{i}'])
        y -= 15

    y= 730
    can.showPage()  # Finaliza a página 2 e começa uma nova página
    can.showPage()  # Finaliza a página 3 e começa uma nova página
    for i in range(1,7):
        if textos[f'descricao_compet_tecnica_0{i}']:
            can.drawString(35, y, " • "+textos[f'descricao_compet_tecnica_0{i}'])
        y -= 15

    match  textos['Escolaridade']:
        case 'Nível Fundamental':
            can.drawString(112.2, 525, '\n')
        case 'Nível Médio':
            can.drawString(210.3, 525, '\n')
        case 'Nível Superior':
            can.drawString(345, 525, '\n')
        case 'Nível Pós-graduação':
            can.drawString(424, 525, '\n')
    can.drawString(328, 470, textos['habilidades_especiais'])
    can.setFont("Helvetica", 9)
    can.drawString(177, 419, textos['experiencia'])


    

    can.save()
    write_pdf(existing_pdf, packet, output_path)


df = pd.read_excel('db.xlsx')
for index, row in df.iterrows():
    textos = row.to_dict()
    fill_pdf_template(textos=textos, output_path=f'{row["cargo"]}_{row["nome"]}.pdf')