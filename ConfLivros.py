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
                texto = leitor.pages[paginas].extract_text()
                return texto if texto is not None else ""
            else:
                textos = []
                for p in paginas:
                    texto = leitor.pages[p].extract_text()
                    textos.append(texto if texto is not None else "")
                return " ".join(textos)
        else:
            textos = []
            for pagina in leitor.pages:
                texto = pagina.extract_text()
                textos.append(texto if texto is not None else "")
            return textos

def extrair_informacoes(texto_primeira, texto_segunda, texto_terceira=None, buscar_administrador=False):
    """Extrai informações do texto fornecido"""
    info = {
        'nome_fundo': None,
        'administrador': None,
        'cnpj': None,
        'periodo': None
    }

    # Extrair CNPJ (continua sendo buscado na primeira página)
    padrao_cnpj = re.compile(r'CNPJ\s*(.{21})', re.IGNORECASE)
    cnpj_match = padrao_cnpj.search(texto_primeira)
    if cnpj_match:
        info['cnpj'] = cnpj_match.group(1).strip()

    # Extrair Nome do Fundo (procurar na segunda página)
    nome_fundo = None
    linhas_segunda = [linha.strip() for linha in texto_segunda.split('\n') if linha.strip()]
    for idx, linha in enumerate(linhas_segunda):
        if re.search(r'Demonstração (?:Financeira|Contábil)', linha, re.IGNORECASE):
            if idx + 2 < len(linhas_segunda):
                nome_fundo = f"{linhas_segunda[idx + 1]} {linhas_segunda[idx + 2]}".strip()
            elif idx + 1 < len(linhas_segunda):
                nome_fundo = linhas_segunda[idx + 1].strip()
            break

    if nome_fundo and "Fundo de Investimento" in nome_fundo:
        info['nome_fundo'] = nome_fundo
    else:
        # Fallback: procurar nas primeiras linhas da segunda página
        for linha in linhas_segunda:
            if "fundo de investimento" in linha.lower():
                info['nome_fundo'] = linha
                break
        if not info['nome_fundo'] and linhas_segunda:
            info['nome_fundo'] = linhas_segunda[0]

    # Extrair Período (procurar na segunda ou terceira página)
    padrao_periodo = re.compile(r'(Referentes ao Exercício Findo em .*?)\n', re.IGNORECASE)
    periodo_match = padrao_periodo.search(texto_segunda)
    if not periodo_match and texto_terceira:
        periodo_match = padrao_periodo.search(texto_terceira)

    # Novo padrão: buscar "diversificação da carteira em"
    if not periodo_match:
        padrao_diversificacao = re.compile(r'30 de\s*(.*?)\n', re.IGNORECASE)
        periodo_match = padrao_diversificacao.search(texto_segunda)
        if not periodo_match and texto_terceira:
            periodo_match = padrao_diversificacao.search(texto_terceira)

    if periodo_match:
        info['periodo'] = periodo_match.group(1).strip()

    # Extrair Administrador (procurar em ambas as páginas)
    if buscar_administrador:
        padrao_admin = re.compile(r'Administrado pel[ao]\s*(.*?)\n', re.IGNORECASE)

        admin_match_primeira = padrao_admin.search(texto_primeira)
        admin_match_segunda = padrao_admin.search(texto_segunda)

        if admin_match_primeira:
            info['administrador'] = admin_match_primeira.group(1).strip()
        elif admin_match_segunda:
            info['administrador'] = admin_match_segunda.group(1).strip()
        else:
            # Busca alternativa nas linhas
            for linha in texto_primeira.split('\n') + texto_segunda.split('\n'):
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
        # Extrair texto das páginas iniciais
        primeira_pagina = extrair_texto_pdf(caminho, 0)
        segunda_pagina = extrair_texto_pdf(caminho, 1) if os.path.getsize(caminho) > 10000 else ""  # Verifica se o arquivo não está vazio
        
        # Extrair informações básicas
        info = extrair_informacoes(primeira_pagina, segunda_pagina, buscar_administrador=True)
        
        # Extrair informações das últimas páginas
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