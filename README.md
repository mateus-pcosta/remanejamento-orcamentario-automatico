# Sistema de Remanejamento Orçamentário Automatizado

Sistema desenvolvido em **Python + Streamlit** para processamento automatizado de planilhas orçamentárias e cálculo de remanejamentos internos e externos entre Unidades Gestoras (UGs).

## Funcionalidades

- Upload de planilhas Excel (.xlsx ou .xls)
- Identificação automática de UGs e naturezas de despesa
- Detecção de déficits orçamentários por natureza
- Configuração de fontes e naturezas proibidas via interface
- Cálculo automático de remanejamento interno (dentro da mesma UG)
- Cálculo automático de remanejamento externo (entre UGs diferentes)
- Validação de resultados (sem saldos negativos)
- Geração de Excel com duas abas:
  - **Saldos Ajustados**: estrutura original com valores corrigidos
  - **Remanejamentos**: detalhamento de todas as transferências
- Geração de arquivo de lote SIAFE para importação direta no sistema
  - Formato de 20 colunas conforme template oficial
  - Preenchimento automático de Tipo de Crédito, Origem de Recursos, UG Emitente, Órgão, Unidade Orçamentária, Programa de Trabalho e Plano Orçamentário
  - Cada remanejamento gera 2 linhas: Redução (Tipo Item 2) + Acréscimo (Tipo Item 1)
  - Regras de mapeamento (Regra 41 e Regra 100) carregadas automaticamente

## Requisitos

- Python 3.8+
- pip

## Instalação

```bash
# Clone o repositório
git clone https://github.com/seu-usuario/remanejamento-orcamentario.git
cd remanejamento-orcamentario

# Instale as dependências
pip install -r requirements.txt
```

## Uso

```bash
streamlit run app.py
```

A aplicação será aberta em `http://localhost:8501`

## Como Funciona

### 1. Upload da Planilha
Faça upload de uma planilha Excel contendo:
- **Coluna A**: Código da fonte (ex: 500, 501)
- **Coluna B**: UG (6 dígitos + nome em MAIÚSCULAS) ou Natureza (6 dígitos + nome)
- **Coluna com "7- Previsão Orçamentária"**: Saldo a ser processado

### 2. Configuração
Configure na interface:
- **Fonte proibida**: código da fonte que não deve participar de remanejamentos
- **Naturezas proibidas**: códigos de naturezas que não devem ser remanejadas

### 3. Processamento
O sistema realiza:
1. Identificação automática da coluna de saldo
2. Remanejamento interno (prioridade)
3. Remanejamento externo (se necessário)
4. Validação dos resultados

### 4. Download
Baixe a planilha ajustada com todos os remanejamentos documentados.

### 5. Geração do Arquivo SIAFE (Lote)
Após o processamento, na seção **Gerar Arquivo SIAFE**:
1. Informe a **Data de Emissão** e o **Número do Processo**
2. Opcionalmente, edite o campo **Observação**
3. Clique em **Gerar Arquivo SIAFE**
4. Baixe o arquivo `.xlsx` pronto para importação no SIAFE-PI

O sistema preenche automaticamente:
- **UG Emitente**: igual à UG Acrescida (UG Destino) de cada linha
- **Tipo de Crédito**: `5` (Remanejamento Interno) para mesma UG; `1` (Suplementar) para UGs diferentes
- **Origem de Recursos**: `0` (Não aplicável) para Tipo 5; `3` (Redução/Anulação) para Tipo 1
- **Órgão, Unidade Orçamentária, Programa de Trabalho, Plano Orçamentário**: derivados da Regra de Mapeamento 41
- **Fonte**: formatada conforme padrão SIAFE (ex: `500` → `5.00`, `501` → `5.01`)
- **Natureza**: formatada conforme padrão SIAFE (ex: `319011` → `3.1.90.11`)

## Regras de Remanejamento

### Prioridades
1. **Remanejamento Interno**: transferências dentro da mesma UG
2. **Remanejamento Externo**: transferências entre UGs da mesma fonte

### Proteções
- Preserva 20% do saldo original de cada natureza doadora
- Limita doações a 40% do saldo por operação
- Prioriza doações únicas para reduzir quantidade de transferências
- Consolida transferências idênticas automaticamente

## Estrutura do Projeto

```
├── app.py                          # Interface Streamlit
├── src/
│   ├── __init__.py
│   ├── processador_orcamento.py    # Lógica de processamento e cálculo de remanejamentos
│   └── gerador_lote.py             # Geração do arquivo de lote SIAFE
├── assets/
│   ├── Itens da Regra de Mapeamento 41.xls   # UG → PT, Órgão, Unidade, Plano
│   ├── Itens da Regra de Mapeamento 100.xls  # Fonte → formato SIAFE
│   └── template_Layout SC_*.xls                  # Template oficial SIAFE (20 colunas)
├── requirements.txt
├── .gitignore
└── README.md
```

## Formato da Planilha de Entrada

A planilha deve seguir o padrão:

| Coluna A (Fonte) | Coluna B (UG/Natureza) | ... | 7- Previsão Orçamentária |
|------------------|------------------------|-----|--------------------------|
| 500 | 140102 - NOME DA UG | ... | 1.000.000,00 |
|     | 319011 - Vencimentos... | ... | -50.000,00 |
|     | 319013 - Obrigações... | ... | 100.000,00 |

- **UGs**: 6 dígitos + " - " + NOME EM MAIÚSCULAS
- **Naturezas**: 6 dígitos + " - " + Nome com minúsculas

## Tecnologias

- [Python 3.8+](https://www.python.org/)
- [Streamlit](https://streamlit.io/)
- [Pandas](https://pandas.pydata.org/)
- [OpenPyXL](https://openpyxl.readthedocs.io/)

## Licença

MIT License - veja o arquivo [LICENSE](LICENSE) para detalhes.

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.
