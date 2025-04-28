import os
import re
import csv
from PyPDF2 import PdfReader
from tkinter import Tk, filedialog

def selecionar_pasta():
    """Permite ao usuário selecionar a pasta com os arquivos PDF"""
    root = Tk()
    root.withdraw()
    pasta = filedialog.askdirectory(title="Selecione a pasta com as demonstrações financeiras")
    return pasta

def extrair_texto_pdf(caminho_pdf, paginas=None):
    """Extrai texto de um PDF, podendo especificar páginas específicas ou todas"""
    with open(caminho_pdf, 'rb') as arquivo:
        leitor = PdfReader(arquivo)
        if paginas is not None:
            if isinstance(paginas, int):
                return leitor.pages[paginas].extract_text()
            else:
                return " ".join([leitor.pages[p].extract_text() for p in paginas])
        else:
            return [pagina.extract_text() for pagina in leitor.pages]

def extrair_informacoes(texto, buscar_administrador=False):
    """Extrai informações do texto fornecido"""
    info = {
        'nome_fundo': None,
        'administrador': None,
        'cnpj': None,
        'periodo': None
    }
    
    # Padrões para extração
    padrao_cnpj = re.compile(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}')
    padrao_periodo = re.compile(r'Período:\s*(\d{2}/\d{2}/\d{4}\s*a\s*\d{2}/\d{2}/\d{4})', re.IGNORECASE)
    padrao_nome_fundo = re.compile(r'Demonstração (?:Financeira|Contábil)\s*(.*?)\n', re.IGNORECASE)
    
    # Extrair CNPJ
    cnpj_match = padrao_cnpj.search(texto)
    if cnpj_match:
        info['cnpj'] = cnpj_match.group()
    
    # Extrair Período
    periodo_match = padrao_periodo.search(texto)
    if periodo_match:
        info['periodo'] = periodo_match.group(1)
    
    # Extrair Nome do Fundo
    nome_match = padrao_nome_fundo.search(texto)
    if nome_match:
        info['nome_fundo'] = nome_match.group(1).strip()
    else:
        # Fallback: pega a primeira linha não vazia
        linhas = [linha.strip() for linha in texto.split('\n') if linha.strip()]
        if linhas:
            info['nome_fundo'] = linhas[0]
    
    # Extrair Administrador se solicitado
    if buscar_administrador:
        padrao_admin = re.compile(r'Administrador(?:a| do Fundo)?:\s*(.*?)(?:\n|$)', re.IGNORECASE)
        admin_match = padrao_admin.search(texto)
        if admin_match:
            info['administrador'] = admin_match.group(1).strip()
        else:
            # Tenta encontrar em linhas que contenham "administrador"
            for linha in texto.split('\n'):
                if 'administrador' in linha.lower() and ':' in linha:
                    info['administrador'] = linha.split(':')[-1].strip()
                    break
    
    return info

def extrair_responsaveis(texto):
    """Extrai contador/diretor responsável do texto"""
    responsaveis = {
        'contador': None,
        'diretor': None
    }
    
    # Padrões melhorados para encontrar os responsáveis
    padroes = [
        (r'(?:Nome do )?Contador(?:\s*Responsável)?:\s*(.+?)(?:\n|$)', 'contador'),
        (r'(?:Nome do )?Diretor(?:\s*Responsável)?:\s*(.+?)(?:\n|$)', 'diretor'),
        (r'Responsável(?: Técnico)?:\s*(.+?)(?:\n|$)', 'contador'),  # Pode capturar ambos
        (r'Assinatura do Contador:\s*(.+?)(?:\n|$)', 'contador'),
        (r'Assinatura do Diretor:\s*(.+?)(?:\n|$)', 'diretor')
    ]
    
    for padrao, campo in padroes:
        matches = re.finditer(padrao, texto, re.IGNORECASE)
        for match in matches:
            valor = match.group(1).strip()
            if valor and not responsaveis[campo]:
                responsaveis[campo] = valor
    
    return responsaveis

def processar_arquivo_pdf(caminho):
    """Processa um único arquivo PDF"""
    try:
        # Extrair informações das primeiras páginas (1-3 para administrador)
        paginas_iniciais = extrair_texto_pdf(caminho, paginas=[0, 1, 2])
        info = extrair_informacoes(paginas_iniciais, buscar_administrador=True)
        
        # Se administrador não foi encontrado nas 3 primeiras páginas, tenta apenas na primeira
        if not info['administrador']:
            primeira_pagina = extrair_texto_pdf(caminho, 0)
            info_primeira = extrair_informacoes(primeira_pagina)
            info.update(info_primeira)
        
        # Extrair informações das últimas páginas (penúltima e última)
        todas_paginas = extrair_texto_pdf(caminho)
        num_paginas = len(todas_paginas)
        ultimas_paginas = []
        
        if num_paginas >= 2:
            ultimas_paginas.append(todas_paginas[-2])  # Penúltima
        if num_paginas >= 1:
            ultimas_paginas.append(todas_paginas[-1])  # Última
        
        texto_ultimas = " ".join(ultimas_paginas)
        responsaveis = extrair_responsaveis(texto_ultimas)
        
        # Combinar resultados
        resultado = {
            'arquivo': os.path.basename(caminho),
            'nome_fundo': info.get('nome_fundo'),
            'administrador': info.get('administrador'),
            'contador': responsaveis.get('contador'),
            'diretor': responsaveis.get('diretor'),
            'cnpj': info.get('cnpj'),
            'periodo': info.get('periodo')
        }
        
        return resultado
    
    except Exception as e:
        print(f"Erro ao processar {os.path.basename(caminho)}: {str(e)}")
        return None

def processar_arquivos_pdf(pasta):
    """Processa todos os arquivos PDF na pasta selecionada"""
    resultados = []
    
    for arquivo in os.listdir(pasta):
        if arquivo.lower().endswith('.pdf'):
            caminho = os.path.join(pasta, arquivo)
            print(f"Processando: {arquivo}")
            
            resultado = processar_arquivo_pdf(caminho)
            if resultado:
                resultados.append(resultado)
    
    return resultados

def salvar_csv(resultados, pasta_saida):
    """Salva os resultados em um arquivo CSV"""
    if not resultados:
        print("Nenhum resultado válido para salvar.")
        return
    
    caminho_csv = os.path.join(pasta_saida, 'resultado_conferencia.csv')
    campos = ['arquivo', 'nome_fundo', 'administrador', 'contador', 'diretor', 'cnpj', 'periodo']
    
    with open(caminho_csv, 'w', newline='', encoding='utf-8') as arquivo_csv:
        escritor = csv.DictWriter(arquivo_csv, fieldnames=campos)
        escritor.writeheader()
        escritor.writerows(resultados)
    
    print(f"Arquivo CSV gerado com sucesso: {caminho_csv}")
    print(f"Total de arquivos processados: {len(resultados)}")

def main():
    print("Selecione a pasta contendo os arquivos PDF das demonstrações financeiras")
    pasta_pdf = selecionar_pasta()
    
    if not pasta_pdf:
        print("Nenhuma pasta selecionada. Operação cancelada.")
        return
    
    resultados = processar_arquivos_pdf(pasta_pdf)
    salvar_csv(resultados, pasta_pdf)

if __name__ == "__main__":
    main()
