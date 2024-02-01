from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter , A4
import io
import textwrap

import pandas as pd

def fill_pdf_template(textos, output_path):
    existing_pdf = PdfReader(open("model_descritivo_cargo.pdf", "rb"))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)


    for textos in textos:
        can.setFont("Helvetica", 11)
        nome = str(textos['nome']) if pd.notna(textos['nome']) else ''
        setor = str(textos['setor']) if pd.notna(textos['setor']) else ''
        cargo = str(textos['cargo']) if pd.notna(textos['cargo']) else ''
        superior_imediato = str(textos['superior_imediato']) if pd.notna(textos['superior_imediato']) else ''
        match textos['regime_contratacao']:
            case 'Comissionado':
                can.drawString(324.6, 584, '\n')
            case 'Seletivo':
                can.drawString(424, 584, '\n')
            case 'Efetivo':
                can.drawString(494, 584, '\n')
        local_trabalho = str(textos['local_trabalho']) if pd.notna(textos['local_trabalho']) else ''
        ch = str(textos['ch']) if pd.notna(textos['ch']) else ''
        objetivo_cargo = str(textos['objetivo_cargo']) if pd.notna(textos['objetivo_cargo']) else ''
        descricao_sumaria_01 = str("• "+textos['descricao_sumaria_01']) if pd.notna(textos['descricao_sumaria_01']) else ''
        descricao_sumaria_02 = str("• "+textos['descricao_sumaria_02']) if pd.notna(textos['descricao_sumaria_02']) else ''
        descricao_sumaria_03 = str("• "+textos['descricao_sumaria_03']) if pd.notna(textos['descricao_sumaria_03']) else ''
        descricao_sumaria_04 = str("• "+textos['descricao_sumaria_04']) if pd.notna(textos['descricao_sumaria_04']) else ''
        descricao_sumaria_05 = str("• "+textos['descricao_sumaria_05']) if pd.notna(textos['descricao_sumaria_05']) else ''
        descricao_sumaria_06 = str("• "+textos['descricao_sumaria_06']) if pd.notna(textos['descricao_sumaria_06']) else ''
        descricao_sumaria_07 = str("• "+textos['descricao_sumaria_07']) if pd.notna(textos['descricao_sumaria_07']) else ''
        can.drawString(150, 665, nome)
        can.drawString(280, 630, setor)
        can.drawString(37, 614, cargo)
        can.drawString(43, 574, superior_imediato)
        can.setFont("Helvetica", 8)
        can.drawString(120, 552, local_trabalho)
        can.setFont("Helvetica", 11)
        can.drawString(362, 550, ch)
        can.drawString(34, 519, objetivo_cargo[:98])
        can.drawString(34, 506, objetivo_cargo[98:194])
        can.drawString(34, 493, objetivo_cargo[194:291])
        can.drawString(34, 480, objetivo_cargo[291:])
        can.drawString(35, 320, descricao_sumaria_01)
        can.drawString(35, 305, descricao_sumaria_02)
        can.drawString(35, 290, descricao_sumaria_03)
        can.drawString(35, 275, descricao_sumaria_04)
        can.drawString(35, 260, descricao_sumaria_05)
        can.drawString(35, 245,descricao_sumaria_06)
        can.drawString(35, 230, descricao_sumaria_07)
        

        can.save()

        packet.seek(0)
        new_pdf = PdfReader(packet)
        page = existing_pdf.pages[0]
        page.merge_page(new_pdf.pages[0])
        output = PdfWriter()
        output.add_page(page)
        with open(output_path, "wb") as outputStream:
            output.write(outputStream)


df = pd.read_excel('db.xlsx')
for index, row in df.iterrows():
    textos=[{'nome':row['nome'],'setor':row['setor'], 'cargo':row['cargo'], 'superior_imediato':row['superior_imediato'],
            'regime_contratacao': row['regime_contratacao'], 'local_trabalho':row['local_trabalho'], 'ch': row['ch'],
            'objetivo_cargo': row['objetivo_cargo'], 'descricao_sumaria_01': row['descricao_sumaria_01'], 'descricao_sumaria_02': row['descricao_sumaria_02'],
            'descricao_sumaria_03': row['descricao_sumaria_03'], 'descricao_sumaria_04': row['descricao_sumaria_04'],
            'descricao_sumaria_05': row['descricao_sumaria_05'], 'descricao_sumaria_06': row['descricao_sumaria_06'],
            'descricao_sumaria_07': row['descricao_sumaria_07']}]
    fill_pdf_template(textos=textos,output_path=f'{row["cargo"]}_{row["nome"]}.pdf')




