# Automação de Contratos de Locação

Este projeto automatiza a geração de contratos de locação de imóveis utilizando dados de uma planilha Excel e criando documentos no formato .docx e .pdf. 
Foram utilizados dados fictícios de pessoas como exemplo.
A seguir estão os detalhes sobre o código e sua funcionalidade.

## Instalação e Importação de Bibliotecas

Para utilizar este projeto, é necessário instalar as seguintes bibliotecas:

```bash
pip install pandas
pip install python-docx
pip install docx2pdf
```

As bibliotecas são importadas no início do script:

```python
from docx import Document as dc
import pandas as pd
from docx2pdf import convert
import os
import shutil
```

## Funcionalidades do Código

### 1. Limpeza das Pastas de Destino

O código remove todos os arquivos das pastas `words` e `pdfs` para garantir que os novos documentos sejam gerados sem conflitos:

```python
for folder in ['words','pdfs']:
    for root, dirs, files in os.walk(f'D:\Documentos\Automação_Python\Automação_contrato_final\{folder}'):
        for f in files:
            os.unlink(os.path.join(root, f))
        for d in dirs:
            shutil.rmtree(os.path.join(root, d))
```

### 2. Leitura da Planilha Excel

Os dados fictícios dos contratos são lidos de uma planilha Excel:

```python
info_table = pd.read_excel('D:\Documentos\Automação_Python\Automação_contrato_final\Informacoes.xlsx')
```

### 3. Geração dos Documentos de Contrato

Para cada linha na planilha, um novo documento de contrato é criado, preenchendo os campos com as informações fornecidas:

```python
for line_number in info_table.index:
    doc = dc("D:\Documentos\Automação_Python\Automação_contrato_final\Contrato_locacao_imovel.docx")
    # Extração das informações da planilha
    nome_locador = str(info_table.loc[line_number,'nome_locador'])
    # (continua com a extração de outras informações)
    
    # Dicionário de substituição
    dict_values = {
        "NOME_LOCADOR": nome_locador,
        # (continua com outros campos)
    }

    # Substituição nos parágrafos do documento
    for parag in doc.paragraphs:
        for cod in dict_values:
            value = dict_values[cod]
            parag.text = parag.text.replace(cod, value)

    doc.save(f'words\Contrato - {nome_locador}.docx')
```

### 4. Conversão para PDF

Os documentos .docx gerados são convertidos para o formato .pdf:

```python
for root, dirs, files in os.walk('D:\Documentos\Automação_Python\Automação_contrato_final\words'):
    for file_name in files:
        convert(f"words\{file_name}",f"pdfs\{file_name.replace('docx','pdf')}")
```

## Conclusão

Este projeto facilita a criação automática de contratos de locação personalizados, utilizando informações de uma planilha Excel e gerando documentos em formatos .docx e .pdf, prontos para uso. Com esta automação, é possível economizar tempo e reduzir erros manuais na preparação de documentos legais.
