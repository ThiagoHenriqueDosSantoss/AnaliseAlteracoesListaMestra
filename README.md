# 📊 Analisador de Relatório MPOP218

Este projeto é um script Python que automatiza a análise do relatório **MPOP218** — histórico de alterações da Lista Mestra — exportado do **ERP Senior**. Ele realiza o pré-processamento, filtragem e simplificação das descrições de alterações para facilitar análises e tomadas de decisão.

> ⚠️ Projeto em desenvolvimento.

---

## ⚙️ Funcionalidades

- 📥 Leitura automatizada de arquivo Excel (.xlsx) exportado do ERP Senior
- 🔍 Filtro de dados por:
  - Período (data inicial e final)
  - Grupo da usina
- 🧹 Limpeza de dados:
  - Remoção de colunas e linhas vazias
- ✂️ Simplificação de descrições da coluna `Alteração`
- 📊 Contagem de tipos específicos de alteração:
  - Quantidade prevista alterada
  - Data do lote alterada
  - Fase alterada
  - Inclusão de item
- 📝 Geração automática do nome do arquivo de saída (evita sobrescrita)
- 📤 Exportação dos resultados processados para um novo arquivo `.xlsx`

---

## 📁 Estrutura Esperada do Arquivo de Entrada

- Formato: `.xlsx`
- Deve conter, obrigatoriamente, as colunas:
  - `Data`
  - `Usina`
  - `Alteração`

> ⚠️ As duas primeiras linhas do arquivo são ignoradas (`skiprows=2`), pois são geralmente cabeçalhos duplicados exportados pelo ERP.

---

## 🧠 Lógica de Simplificação

A função `simplificar_descricao()` usa **expressões regulares** para remover o trecho intermediário das descrições, mantendo apenas a parte inicial relevante.

**Exemplo:**

"Quantidade prevista alterada de 100 para 200"
➡️ "Quantidade prevista alterada"


## 📚 Dependências Utilizadas

- `pandas` - 	Manipulação e análise de dados tabulares	- (Como instalar a dependencia) pip install pandas 
- `openpyxl` - 	Leitura e escrita de arquivos Excel .xlsx - (Como instalar a dependencia) 	pip install openpyxl
- `tqdm` - 	Exibição de barra de progresso - (Como instalar a dependencia)	pip install tqdm
