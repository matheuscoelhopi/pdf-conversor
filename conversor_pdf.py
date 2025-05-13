'''
import pdfplumber
import pandas as pd


caminho_pdf = input("Caminho do PDF: ")
# Abrir o PDF
with pdfplumber.open(caminho_pdf) as pdf:
    dados = []

    for pagina in pdf.pages:
        texto = pagina.extract_text()
        if texto:
            linhas = texto.split("\n")
            for linha in linhas:
                # Filtros simples para ignorar cabeçalhos e rodapés
                if linha.strip() and any(char.isdigit() for char in linha[:10]):
                    dados.append(linha)

# Processar os dados extraídos
linhas_limpa = []
for linha in dados:
    partes = linha.split()
    if len(partes) >= 7:
        matricula = partes[0]
        nome = " ".join(partes[1:-4])
        lotacao = partes[-4]
        ordem = partes[-3]
        cargo = " ".join(partes[-2:])
        linhas_limpa.append([ordem, matricula, nome, cargo, cls1, ref, portaria])

# Criar o DataFrame
df = pd.DataFrame(linhas_limpa, columns=[ "Ordem", "Matrícula", "Nome", "Cargo", "Cls", "Ref", "Portaria"])

# Salvar como Excel
df.to_excel("convertido2.xlsx", index=False)
"C:/Users/mathe/Downloads/consultaServidorLotacaoFiltro.jsf;jsessionid=6E45B8AACDB489A51260C87F528B9891.pdf"
'''

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
                # Padrão regex para identificar linhas de dados
                # Exemplo: "1 828825 PAULA RENATA MACHADO CORREA CORREGEDORIA AGENTE DA RECEITA ESTADUAL PORTARIA 50"
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