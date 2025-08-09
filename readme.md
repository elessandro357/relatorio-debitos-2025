# 📊 Relatório de Débitos e Saldos — 2025

Aplicação interativa em **Streamlit** para análise de **débitos** e **saldos bancários** de secretarias, com dashboards, filtros, exportação para **Excel** e **PDF**, e validações automáticas.

---

## 🚀 Funcionalidades
- **Upload** de planilhas (`.csv`, `.xls`, `.xlsx`)
- **Mapeamento de colunas** para arquivos com cabeçalhos diferentes
- **Validação automática**:
  - Datas válidas
  - Valores numéricos ≥ 0
  - CNPJ com 14 dígitos (zeros à esquerda se necessário)
- **Consolidação de duplicados** (opcional)
- **Marcação de outliers** (> percentil 95 por secretaria)
- **Filtros persistentes** e botão **Limpar filtros**
- **Totais dinâmicos** na tela
- **Exportação**:
  - Excel com formatação de moeda BRL e aba de resumo
  - PDF com quebra por secretaria, total por grupo e numeração de páginas
- **Templates** de Débitos e Saldos para download

---

## 📂 Estrutura do Repositório
