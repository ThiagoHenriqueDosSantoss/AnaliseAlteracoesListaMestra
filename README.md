# 📊 Analisador de Relatório

Este script automatiza a análise do relatório **MPOP218** (Histórico de alterações da Lista Mestra), exportado do **ERP Senior**, realizando o pré-processamento e simplificação das descrições de alterações.
(Em desenvolvimento)
---

## ⚙️ Funcionalidades

- 📥 Leitura do arquivo Excel exportado do ERP Senior
- 🧹 Limpeza de dados: remoção de colunas e linhas completamente vazias
- ✂️ Simplificação de descrições contidas na coluna `Alteração`
- 📁 Geração automática do nome do arquivo de saída, evitando sobreposição
- 📤 Exportação do resultado para um novo arquivo `.xlsx`

---

## 📁 Estrutura Esperada do Arquivo de Entrada

O arquivo de entrada deve ser um `.xlsx` com pelo menos uma coluna chamada **"Alteração"**, que conterá os textos a serem simplificados.

> ⚠️ O script ignora as duas primeiras linhas do arquivo ao carregar os dados (`skiprows=2`).

---

## 🧠 Lógica de Simplificação

A função `simplificar_descricao` remove o trecho intermediário das descrições usando uma **expressão regular**, com base na necessidade de extrair a parte desejada.
