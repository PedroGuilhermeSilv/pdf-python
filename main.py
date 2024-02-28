from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import io
import textwrap
import pandas as pd
import os


def create_canvas(textos):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.setFont("Helvetica", 10)
    for key, value in textos.items():
        if pd.notna(value):
            textos[key] = str(value)
        else:
            textos[key] = ""
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
    can.drawString(150, 665, textos["nome"])
    can.drawString(245, 618, textos["setor"])
    if len(textos["cargo"]) > 41:
        parte1 = textos["cargo"][:41]
        parte2 = textos["cargo"][41:]
        can.drawString(33, 624, parte1)
        can.drawString(33, 614, parte2)
    else:
        can.drawString(33, 618, textos["cargo"])

    if len(textos["superior_imediato"]) > 41:
        parte1 = textos["superior_imediato"][:41]
        parte2 = textos["superior_imediato"][41:]
        can.drawString(33, 580, parte1)
        can.drawString(33, 570, parte2)
    else:
        can.drawString(33, 578, textos["superior_imediato"])

    can.setFont("Helvetica", 8.8)
    can.drawString(120, 550, textos["local_trabalho"])
    can.setFont("Helvetica", 9.5)
    can.drawString(357, 550, textos["ch"])
    lines = textwrap.wrap(textos["objetivo_cargo"], width=98)
    match textos["regime_contratacao"]:
        case "Comissionado":
            can.drawString(325.5, 593, "\n")
        case "Seletivo":
            can.drawString(424, 591, "\n")
        case "Efetivo":
            can.drawString(494, 591, "\n")
    y = 508
    for line in lines:
        can.drawString(33, y, line)
        y -= 13
    y = 379
    for i in range(1, 8):
        if textos[f"descricao_sumaria_0{i}"]:
            can.drawString(33, y, " • " + textos[f"descricao_sumaria_0{i}"])
        y -= 15
    y = 180
    can.setFont("Helvetica", 10)
    for i in range(1, 6):
        if textos[f"descricao_compet_comportamentais_0{i}"]:
            can.drawString(
                33, y, " • " + textos[f"descricao_compet_comportamentais_0{i}"]
            )
        y -= 15

    y = 740
    can.showPage()  # Finaliza a página 2 e começa uma nova página
    can.setFont("Helvetica", 10)
    for i in range(1, 7):
        if textos[f"descricao_compet_tecnica_0{i}"]:
            text = " • " + textos[f"descricao_compet_tecnica_0{i}"]
            lines = textwrap.wrap(text, width=110, drop_whitespace=True)
            for line in lines:
                can.drawString(33, y, line)
                y -= 15

    match textos["Escolaridade"]:
        case "Nível Fundamental":
            can.drawString(112, 590.5, "\n")
        case "Nível Médio":
            can.drawString(211, 590.5, "\n")
        case "Nível Superior":
            can.drawString(345, 590.5, "\n")
        case "Nível Técnico":
            can.drawString(274, 590.5, "\n")
        case "Nível Pós-graduação":
            can.drawString(424, 590.5, "\n")

    y = 569
    can.setFont("Helvetica", 10)
    text = textos["habilidades_especiais"].split(";")
    for i in text:
        if i != "":
            y -= 15
            if "Desejável" in i:
                can.drawString(33, y, i)
            else:
                if len(i) > 99:
                    i = i.replace("\n", "")
                    can.drawString(33, y, " • " + i[:101])
                    y -= 15
                    i = i[101:]
                    i = i.replace("\n", "")
                    can.drawString(45, y, i)
                else:
                    i = i.replace("\n", "")
                    can.drawString(33, y, " • " + i)

    text = []
    can.setFont("Helvetica", 8.5)
    can.drawString(171, 410.2, textos["experiencia"])

    can.save()
    write_pdf(existing_pdf, packet, output_path)


df = pd.read_excel("db.xlsx")
nomes = []


for index, row in df.iterrows():
    textos = row.to_dict()
    setor = row["setor"]
    if setor not in nomes:
        nomes.append(setor)
        folder_path = f"{setor}"
        os.makedirs(folder_path, exist_ok=True)
    folder_path = f"{setor}"
    fill_pdf_template(
        textos=textos, output_path=f'{folder_path}/{row["cargo"]}_{row["nome"]}.pdf'
    )
