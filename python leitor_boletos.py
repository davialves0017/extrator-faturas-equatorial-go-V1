import glob
import pdfplumber
import csv
import os

pasta_pdfs = r"C:\Users\Ludmila\Downloads\extrator"
caminho_csv = os.path.join(pasta_pdfs, "itens_fatura_por_colunas.csv")

todas_linhas = []

for caminho_pdf in glob.glob(os.path.join(pasta_pdfs, "2026*Cliente*Equatorial*.pdf")):
    nome_arquivo = os.path.basename(caminho_pdf)
    print(f"Lendo: {nome_arquivo}")

    with pdfplumber.open(caminho_pdf) as pdf:
        texto = ""
        for page in pdf.pages:
            texto += page.extract_text() + "\n"

    # Procura bloco itens de fatura
    inicio = texto.lower().find("itens de fatura")
    if inicio == -1:
        continue

    # Pega próximo bloco de 10 linhas após "itens de fatura"
    linhas = texto[inicio:].split('\n')[:20]
    
    for linha in linhas:
        linha_limpa = re.sub(r'[ ]+', ' ', linha.strip())  # limpa múltiplos espaços
        
        # Se tem números com vírgula, é item da tabela
        if re.search(r'kWh.*[\d,]+\s+[\d,]+', linha_limpa):
            partes = linha_limpa.split()
            
            # Mapeia por posição fixa da tabela
            if len(partes) >= 8:
                todas_linhas.append({
                    "arquivo": nome_arquivo,
                    "descricao": ' '.join(partes[:2]),  # 1º e 2º = descrição
                    "unidade": partes[2],               # 3º = kWh
                    "quantidade": partes[3],            # 4º = 30,00
                    "preco_unit": partes[4],            # 5º = 0,008772
                    "valor_rs": partes[5],              # 6º = 0,26
                    "pis_cofins": partes[6],            # 7º = 0,01
                    "base_icms": partes[7],             # 8º = 0,26
                    "aliquota_icms": partes[8] if len(partes)>8 else "",
                    "icms_rs": partes[9] if len(partes)>9 else ""
                })
                print(f"  → {partes[:6]}")  # mostra o que achou

if todas_linhas:
    with open(caminho_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=todas_linhas[0].keys())
        writer.writeheader()
        writer.writerows(todas_linhas)

    print(f"✅ CSV com colunas separadas: itens_fatura_por_colunas.csv")
else:
    print("❌ Nenhuma linha encontrada")

print("FIM")
