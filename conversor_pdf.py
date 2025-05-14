import pandas as pd
import pdfplumber
import re

def pdf_to_xlsx_structured(pdf_path, xlsx_path):
    # Lista para armazenar os dados
    dados = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            
            # Processa cada linha do texto
            for line in text.split('\n'):

                match = re.match(r'^(\d+)\s+(\d+)\s+(.*?)\s+(.*?)\s+(.*?)\s+(.*?)\s+(.*)$', line.strip())
                
                if match:
                    ordem, matricula, nome, lotacao, cargo, cls_ref, portaria = match.groups()
                    
                    # Separa Cls e Ref quando estão juntos (ex: "A 1", "ESPE 11")
                    cls_ref_split = cls_ref.split()
                    cls = cls_ref_split[0] if len(cls_ref_split) > 0 else ''
                    ref = cls_ref_split[1] if len(cls_ref_split) > 1 else ''
                    
                    # Adiciona à lista de dados
                    dados.append({
                        'Ordem': ordem,
                        'Matrícula': matricula,
                        'Nome': nome,
                        'Cargo': cargo,
                        'Cls': cls,
                        'Ref': ref,
                        'Portaria': portaria
                    })
    
    # Cria DataFrame
    df = pd.DataFrame(dados)
    
    # Salva como XLSX
    df.to_excel(xlsx_path, index=False)
    print(f"Arquivo convertido com sucesso: {xlsx_path}")

# Uso
pdf_path = "C:/Users/mathe/Downloads/consultaServidorLotacaoFiltro.jsf;jsessionid=6E45B8AACDB489A51260C87F528B9891.pdf"
xlsx_path = "servidores.xlsx"
pdf_to_xlsx_structured(pdf_path, xlsx_path)