# 📄 Resumo do Script de Extração de Dados Financeiros

## 📌 Visão Geral
Script Python que automatiza a extração de informações estruturadas de demonstrações financeiras em PDF e exporta para CSV.

## 🛠️ Funcionalidades Principais
- **Seleção interativa** de pasta contendo os PDFs
- **Processamento em lote** de múltiplos arquivos
- **Extração inteligente** de:
  - Nome do fundo
  - CNPJ
  - Período de referência
  - Administrador
  - Contador responsável
  - Diretor responsável
- **Geração automática** de relatório em CSV

## 🔍 Métodos Principais

### `selecionar_pasta()`
- Abre diálogo para seleção do diretório
- Retorna caminho da pasta selecionada

### `extrair_texto_pdf(caminho_pdf, paginas=None)`
- Extrai texto de PDFs
- Permite especificar páginas específicas

### `extrair_informacoes(texto, buscar_administrador=False)`
- Utiliza regex para capturar:
  - Padrões de CNPJ (xx.xxx.xxx/xxxx-xx)
  - Períodos (dd/mm/aaaa a dd/mm/aaaa)
  - Nome do fundo
  - Administrador (quando solicitado)

### `extrair_responsaveis(texto)`
- Identifica contador e diretor responsáveis
- Busca em múltiplos padrões de texto

### `processar_arquivo_pdf(caminho)`
- Orquestra o processamento completo de um arquivo
- Combina dados das primeiras e últimas páginas

## 📊 Saída
Gera arquivo `resultado_conferencia.csv` com estrutura:

| Coluna         | Descrição                     |
|----------------|-------------------------------|
| arquivo        | Nome do arquivo original      |
| nome_fundo     | Nome completo do fundo        |
| administrador  | Entidade administradora       |
| contador       | Contador responsável          |
| diretor        | Diretor responsável           |
| cnpj           | CNPJ do fundo (formatado)     |
| periodo        | Período da demonstração       |

## ⚙️ Requisitos
- Python 3.x
- Bibliotecas:
  - PyPDF2
  - tkinter
  - re (regex)
  - csv

## 🚀 Como Usar
1. Execute o script
2. Selecione a pasta com os PDFs
3. Aguarde o processamento
4. Verifique o arquivo CSV gerado

> **Nota:** O script inclui tratamento de erros e fallbacks para casos onde os padrões esperados não são encontrados.
