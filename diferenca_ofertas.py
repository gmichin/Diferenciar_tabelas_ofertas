import pdfplumber
import pandas as pd
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

def sanitize_sheet_name(name):
    """Remove caracteres inválidos para nomes de aba do Excel"""
    invalid_chars = ['\\', '/', '*', '[', ']', ':', '?']
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name[:31]  # Limita a 31 caracteres

def extract_date_from_pdf(pdf_path):
    """Extrai a data do texto do PDF"""
    date_pattern = r'\d{2}\s?[/-]\s?\d{2}\s?[/-]\s?\d{2,4}|\d{1,2}\s?de\s?[A-Za-zç]+\s?de\s?\d{4}'
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            
            # Procura acima da LEGENDA
            if 'LEGENDA' in text:
                legenda_index = text.index('LEGENDA')
                search_area = text[:legenda_index]
                date_match = re.search(date_pattern, search_area)
                if date_match:
                    return sanitize_sheet_name(date_match.group())
            
            # Procura no canto direito
            right_side = text[-200:]
            date_match = re.search(date_pattern, right_side)
            if date_match:
                return sanitize_sheet_name(date_match.group())
    
    return "Data_nao_encontrada"

def process_pdf_to_dataframe(pdf_path):
    """Processa um único PDF e retorna um DataFrame e a data"""
    filename = os.path.basename(pdf_path).replace('.pdf', '')
    date = extract_date_from_pdf(pdf_path)
    
    # Extrair todas as tabelas
    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                all_tables.extend(table)
    
    if not all_tables:
        print(f"Nenhuma tabela encontrada em {pdf_path}.")
        return None, None, None
    
    # Encontrar cabeçalho - procura a linha com "CÓD. REF"
    headers = ["CÓD. REF", "MARCA", "PRODUTO", "PESO CAIXA", 
              "VALOR 3%", "VALOR 1%", "STATUS", "ESTOQUE", "CX's"]
    
    # Encontrar a linha de cabeçalho real
    header_row_idx = None
    for i, row in enumerate(all_tables):
        if any("CÓD. REF" in str(cell) for cell in row):
            header_row_idx = i
            break
    
    if header_row_idx is None:
        print(f"Cabeçalho não encontrado em {pdf_path}, usando padrão")
        header_row_idx = 0
        actual_headers = all_tables[header_row_idx]
    else:
        actual_headers = all_tables[header_row_idx]
    
    # Coletar todas as linhas abaixo do cabeçalho
    data_rows = []
    for row in all_tables[header_row_idx+1:]:
        if len(row) == len(actual_headers):  # Só adiciona se tiver número correto de colunas
            data_rows.append(row)
    
    # Criar DataFrame
    df = pd.DataFrame(data_rows, columns=actual_headers)
    
    # Limpeza básica - remover linhas vazias
    df = df[df.iloc[:, 0].str.strip().astype(bool)]
    
    return df, date, filename

def pdfs_to_excel_with_sheets(pdf_paths, output_excel_path=None):
    """Converte múltiplos PDFs para um único Excel com abas diferentes"""
    try:
        if not pdf_paths:
            print("Nenhum arquivo PDF fornecido.")
            return
        
        if output_excel_path is None:
            # Usa o diretório do primeiro PDF como base
            output_excel_path = os.path.join(os.path.dirname(pdf_paths[0]), "Tabelas_Consolidadas.xlsx")
        
        # Criar um writer Excel
        writer = pd.ExcelWriter(output_excel_path, engine='openpyxl')
        
        for pdf_path in pdf_paths:
            df, date, filename = process_pdf_to_dataframe(pdf_path)
            if df is not None:
                # Usar a data sanitizada como nome da aba
                sheet_name = date if date else sanitize_sheet_name(filename)
                
                # Se a data já existir como aba, adiciona um sufixo
                original_sheet_name = sheet_name
                counter = 1
                while sheet_name in writer.book.sheetnames:
                    sheet_name = f"{original_sheet_name}_{counter}"
                    counter += 1
                
                # Escrever no Excel
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Formatar o cabeçalho
                ws = writer.book[sheet_name]
                ws.insert_rows(1)
                ws['A1'] = f"{filename} - {date.replace('_', ' ') if date else filename}"
                ws.merge_cells('A1:I1')
                
                # Configurar estilo
                header_font = Font(bold=True, size=12)
                header_alignment = Alignment(horizontal='center')
                
                for row in ws.iter_rows(min_row=1, max_row=1):
                    for cell in row:
                        cell.font = header_font
                        cell.alignment = header_alignment
                
                print(f"Processado: {pdf_path} -> aba '{sheet_name}' ({len(df)} linhas)")
        
        # Salvar o arquivo Excel
        writer.close()
        print(f"\nArquivo Excel gerado com sucesso: {output_excel_path}")
    
    except Exception as e:
        print(f"Erro durante o processamento: {str(e)}")

# Exemplo de uso - processa todos os PDFs no diretório
pdf_directory = r"C:\Users\win11\Downloads\Tabelas de ofertas"
pdf_files = [os.path.join(pdf_directory, f) for f in os.listdir(pdf_directory) if f.endswith('.pdf')]

if pdf_files:
    pdfs_to_excel_with_sheets(pdf_files)
else:
    print("Nenhum arquivo PDF encontrado no diretório especificado.")