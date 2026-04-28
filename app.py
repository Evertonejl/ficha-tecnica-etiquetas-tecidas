from flask import Flask, render_template, request, send_file
import re
import tempfile
import os
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, Image, TableStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.units import mm

app = Flask(__name__)

# ===============================
# LIMPAR NOME DO ARQUIVO
# ===============================
def limpar_nome_arquivo(nome):
    nome = re.sub(r'[\\/*?:"<>|]', "", nome)
    return nome.strip()

# ===============================
# LEITURA TXT
# ===============================
def ler_txt(caminho):
    with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()

# ===============================
# EXTRAÇÃO MUCAD
# ===============================
def extrair_dados(texto):
    dados = {}
    cores = []
    total_metros = 0
    total_picks = 0

    linhas = texto.splitlines()

    for linha in linhas:
        if re.match(r"^\s*\d+\s+\d+\s+0\.[0-9]", linha):
            partes = linha.split()

            if len(partes) >= 5:
                sel = partes[0]
                picks = int(partes[1])
                metros = int(partes[4])

                total_metros += metros
                total_picks += picks

                cores.append({
                    "cor": f"Cor {sel}",
                    "picks": picks,
                    "metros": metros
                })

    dados["cores"] = cores
    dados["total_consumo"] = total_metros
    dados["total_picks"] = total_picks

    return dados

# ===============================
# IMAGEM EM MM
# ===============================
def imagem_em_mm(caminho, largura_mm=None, altura_mm=None):
    img = Image(caminho)

    if largura_mm and altura_mm:
        img.drawWidth = largura_mm * mm
        img.drawHeight = altura_mm * mm
    else:
        img.drawWidth = img.imageWidth
        img.drawHeight = img.imageHeight

    return img

# ===============================
# PDF
# ===============================
def gerar_pdf(dados, nome, desenho, batida, caminho_img=None, largura_mm=None, altura_mm=None):
    doc = SimpleDocTemplate(
        nome,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=0
    )

    styles = getSampleStyleSheet()
    e = []

    e.append(Paragraph("FICHA TÉCNICA DE CONSUMO", styles["Title"]))
    e.append(Spacer(1, 15))

    data = datetime.now().strftime("%d/%m/%Y")
    e.append(Paragraph(f"Data: {data}", styles["Normal"]))
    e.append(Spacer(1, 10))

    if desenho:
        e.append(Paragraph(f"<b>Nome do Desenho:</b> {desenho}", styles["Normal"]))

    if batida:
        e.append(Paragraph(f"<b>Batida:</b> {batida}", styles["Normal"]))

    e.append(Spacer(1, 20))

    e.append(Paragraph("Dados de Consumo", styles["Heading2"]))
    e.append(Spacer(1, 10))

    tabela = [["Cor", "Consumo-cor", "Consumo (m)"]]

    for c in dados["cores"]:
        tabela.append([c["cor"], str(c["picks"]), str(c["metros"])])

    tabela.append([
        "TOTAL",
        str(dados["total_picks"]),
        str(dados["total_consumo"])
    ])

    t = Table(tabela, colWidths=[150, 150, 150])

    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1f2937")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
    ]))

    e.append(t)
    e.append(Spacer(1, 25))

    if caminho_img and os.path.exists(caminho_img):
        e.append(Paragraph("Imagem da Etiqueta", styles["Heading2"]))
        e.append(Spacer(1, 10))

        img = imagem_em_mm(caminho_img, largura_mm, altura_mm)
        img.hAlign = 'CENTER'

        e.append(img)
        e.append(Spacer(1, 30))

    e.append(Paragraph(
        "Relatório técnico gerado automaticamente a partir do MuCAD.",
        styles["Italic"]
    ))

    doc.build(e)

# ===============================
# ROTAS
# ===============================
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    # TXT — salvo em arquivo temporário
    arquivo = request.files["arquivo"]
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as tmp_txt:
        arquivo.save(tmp_txt.name)
        caminho_txt = tmp_txt.name

    texto = ler_txt(caminho_txt)
    dados = extrair_dados(texto)
    os.unlink(caminho_txt)  # limpa o arquivo temporário

    # CAMPOS
    desenho = request.form.get("desenho", "")
    batida = request.form.get("batida", "")

    # NOME DO PDF
    nome_pdf = request.form.get("nome_pdf", "ficha_tecnica")
    nome_pdf = limpar_nome_arquivo(nome_pdf)
    if not nome_pdf.lower().endswith(".pdf"):
        nome_pdf += ".pdf"

    # TAMANHO DA IMAGEM
    largura_mm = request.form.get("largura_mm")
    altura_mm = request.form.get("altura_mm")
    largura_mm = float(largura_mm) if largura_mm else None
    altura_mm = float(altura_mm) if altura_mm else None

    # IMAGEM — salva em arquivo temporário
    imagem = request.files.get("imagem")
    caminho_img = None
    if imagem and imagem.filename != "":
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
            imagem.save(tmp_img.name)
            caminho_img = tmp_img.name

    # GERAR PDF em arquivo temporário
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        caminho_pdf = tmp_pdf.name

    gerar_pdf(dados, caminho_pdf, desenho, batida, caminho_img, largura_mm, altura_mm)

    # Limpa imagem temporária se existir
    if caminho_img and os.path.exists(caminho_img):
        os.unlink(caminho_img)

    return send_file(caminho_pdf, as_attachment=True, download_name=nome_pdf)

# ===============================
# MAIN
# ===============================
if __name__ == "__main__":
    app.run(debug=False)