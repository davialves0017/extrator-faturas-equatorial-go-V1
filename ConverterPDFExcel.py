import pdfplumber
import pandas as pd
import re
import os

def final_extractor(pdf_path):
    """
    Extrai os itens da fatura usando padrões Regex específicos e cirúrgicos,
    construídos a partir da análise do texto bruto.
    """
    all_items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                # Usar layout=True pode ajudar a preservar o espaçamento original
                text = page.extract_text(layout=True)
                if text:
                    full_text += text + "\n"
            
            if not full_text:
                return []

            # 1. Extrair Metadados (UC, Data, etc.)
            uc_match = re.search(r'(\d{8,10})', full_text) # Pega o primeiro número longo que provavelmente é a UC
            mes_ano_match = re.search(r'([A-Z]{3}/\d{4})', full_text)
            
            uc = uc_match.group(1) if uc_match else "N/A"
            mes_ano = mes_ano_match.group(1) if mes_ano_match else "N/A"

            # 2. Definir os Padrões de Extração (Regex)
            
            # Padrão para itens principais (CONSUMO, BANDEIRA, INJEÇÃO)
            # Normaliza espaços múltiplos para um único espaço para facilitar a correspondência
            normalized_text = re.sub(r'\s+', ' ', full_text)
            
            # Regex para encontrar os itens que começam com uma descrição conhecida
            main_item_pattern = re.compile(
                r"(FORNECIMENTO|ADC BANDEIRA AMARELA|ADC BANDEIRA VERMELHA|CONSUMO N.O COMPENSADO|CONSUMO SCEE|INJEÇÃO SCEE - UC \d+ - GD I|ENERGIA COMP N.O ISENTA \(TRIBUTOS\) - UC)" # Descrições
                r"\s+(kWh)?\s*" # Unidade (opcional)
                r"([\d\.,-]+)\s+" # Quant.
                r"([\d\.,-]+)\s+" # Preço Unit.
                r"([\d\.,-]+)\s+" # Valor
                r"([\d\.,-]+)\s+" # PIS/COFINS (valor)
                r"([\d\.,-]+)\s+" # Base ICMS
                r"([\d\.,%]+)\s+" # Alíquota ICMS
                r"([\d\.,-]+)\s*" # ICMS (valor)
                r"([\d\.,-]+)?"  # Tarifa Unit. (opcional)
            )

            # Padrão para itens simples (CONTRIBUIÇÃO, DEVOLUÇÃO, etc.)
            simple_item_pattern = re.compile(
                r"(CONTRIB\. ILUM\. P\.BLICA - MUNICIPAL|DEV\. VAL\. COBR\. A MAIOR \(-\))\s+([\d\.,-]+)"
            )
            
            # Padrão para os impostos que aparecem na lateral
            tax_pattern = re.compile(r"(PIS/PASEP|ICMS|COFINS)\s+([\d\.,]+)\s+([\d\.,%]+)\s+([\d\.,-]+)")

            # 3. Encontrar todas as correspondências no texto normalizado
            
            # Itens principais
            for match in main_item_pattern.finditer(normalized_text):
                all_items.append({
                    'Descrição': match.group(1).strip(),
                    'Unid.': match.group(2) if match.group(2) else 'kWh',
                    'Quant.': match.group(3),
                    'Preço unit (R$)': match.group(4),
                    'Valor (R$)': match.group(5),
                    'PIS/COFINS (val)': match.group(6),
                    'Base Calc. ICMS (R$)': match.group(7),
                    'Alíq. ICMS (%)': match.group(8),
                    'ICMS (R$)': match.group(9),
                    'Tarifa unit. (R$)': match.group(10) or '0,00'
                })

            # Itens simples
            for match in simple_item_pattern.finditer(normalized_text):
                all_items.append({'Descrição': match.group(1).strip(), 'Valor (R$)': match.group(2)})
            
            # Impostos laterais
            for match in tax_pattern.finditer(normalized_text):
                all_items.append({
                    'Descrição': match.group(1),
                    'Base de Cálculo': match.group(2),
                    'Alíq. (%)': match.group(3),
                    'Valor (R$)': match.group(4)
                })

            # Adicionar metadados a todos os itens encontrados
            for item in all_items:
                item['Arquivo'] = os.path.basename(pdf_path)
                item['UC'] = uc
                item['Mês/Ano'] = mes_ano
                
    except Exception as e:
        print(f"  -> ERRO CRÍTICO no arquivo {os.path.basename(pdf_path)}: {e}")
    
    return all_items

def main():
    # CAMINHO ATUALIZADO AQUI
    pdf_directory = r'C:\Users\Ludmila Evangelista\Documents\extrator'
    
    output_file = 'RELATORIO_FINAL_FATURAS.xlsx'
    
    if not os.path.exists(pdf_directory):
        print(f"ERRO: Diretório não encontrado: {pdf_directory}")
        return

    pdf_files = [f for f in os.listdir(pdf_directory) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print(f"AVISO: Nenhum arquivo PDF encontrado em {pdf_directory}")
        return

    print(f"Iniciando extração final em {len(pdf_files)} arquivo(s)...")
    all_data = []
    for filename in pdf_files:
        pdf_path = os.path.join(pdf_directory, filename)
        print(f"Processando: {filename}")
        # Normalizando o texto antes de passar para o extrator
        items = final_extractor(pdf_path)
        if items:
            all_data.extend(items)
            print(f"  -> Sucesso! {len(items)} itens extraídos.")
        else:
            print("  -> Aviso: Nenhum item correspondente encontrado.")

    if not all_data:
        print("\nProcessamento concluído, mas nenhum dado foi extraído de nenhum arquivo.")
        return

    df = pd.DataFrame(all_data)
    
    # Organizar colunas na ordem desejada
    column_order = [
        'Arquivo', 'UC', 'Mês/Ano', 'Descrição', 'Unid.', 'Quant.', 
        'Preço unit (R$)', 'Valor (R$)', 'PIS/COFINS (val)', 
        'Base Calc. ICMS (R$)', 'Alíq. ICMS (%)', 'ICMS (R$)', 'Tarifa unit. (R$)',
        'Base de Cálculo', 'Alíq. (%)'
    ]
    df = df.reindex(columns=column_order)

    # Salvar em Excel com formatação
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Itens da Fatura', index=False)
        workbook  = writer.book
        worksheet = writer.sheets['Itens da Fatura']
        
        # Formato do cabeçalho
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Auto-ajuste da largura das colunas
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col))
            worksheet.set_column(i, i, column_len + 2)

    print(f"\nSUCESSO! Relatório final gerado: {output_file}")

if __name__ == "__main__":
    main()