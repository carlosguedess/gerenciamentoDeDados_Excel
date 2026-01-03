# verifica_enderecos.py — README

✅ **Visão geral**

Este script lê um arquivo Excel chamado `enderecos.xlsx`, normaliza e analisa endereços residenciais, identifica os endereços mais repetidos e gera um arquivo Excel de saída `clientes_enderecos_<YYYY-MM-DD>.xlsx` com duas abas:
- `Todos_Enderecos`: todos os dados lidos e algumas colunas auxiliares (ex.: `Qtd_Repeticoes`, `Endereco_Repetido`).
- `Resumo`: uma linha por cliente para os endereços mais frequentes (TOP N), com informações como `Cliente`, `CodigoRevendedor`, `Bloqueado` e `Estrutura` (se presentes).

---

## 🔧 Requisitos
- Python 3.8+
- pandas
- openpyxl (recomendado para gravar arquivos Excel e, se desejar, ocultar abas)

Instalação rápida:

```bash
pip install pandas openpyxl
```

---

## ⚙️ Entradas esperadas
- Nome do arquivo: `enderecos.xlsx` (coloque na mesma pasta do script ou ajuste o caminho no código).
- Colunas esperadas (algumas são detectadas automaticamente com várias variações):
  - Coluna de endereço: `RuaResidencial` (obrigatória para normalizar e identificar duplicados)
  - Coluna de cliente: `Cliente`, `Nome`, `nome_cliente`, etc. (o script tenta detectar automaticamente)
  - Coluna de código/ID: `Codigo`, `Código`, `ID`, `ClienteID`, `CodigoRevendedor`, etc.
  - Coluna de bloqueio (opcional): `Bloqueado`, `Bloqueio`, `blocked`, etc. — o valor será incluído no `Resumo` quando existir.
  - Coluna de estrutura (opcional): `Estrutura`, `Estrutra`, `setor`, `local`, etc. — usada para filtrar o `Resumo` para registros relacionados a *central de inicios* (quando detectada).

> Observação: o código é tolerante a variações de nomes (case, underscores, espaços). Se sua coluna tiver um nome muito diferente, informe-me que eu adiciono à lista de detecção.

---

## 🧠 O que o script faz (passo a passo)
1. Lê `enderecos.xlsx` com `pd.read_excel`.
2. Normaliza a coluna `RuaResidencial` (strip + lower) para evitar diferenças por caixa e espaços.
3. Marca duplicados na coluna `Endereco_Repetido` ("Duplicado" / "Único").
4. Conta quantas vezes cada `RuaResidencial` aparece (`contagem`) e pega os TOP_N (variável `TOP_N`, padrão no script atual).
5. Detecta automaticamente colunas:
   - nome do cliente (`name_col`)
   - código/ID do cliente (`code_col`)
   - coluna de bloqueio (`block_col`) — se presente, seu valor é copiado para o `Resumo`
   - coluna de estrutura (`estrutura_col`) — se presente, seu valor é copiado para o `Resumo` e usado no filtro
6. Para cada endereço do TOP_N, o script cria uma linha por cliente contendo: `Endereco`, `Quantidade`, `Cliente`, `CodigoRevendedor`, `Bloqueado` e `Estrutura` (se existirem).
7. Constrói o DataFrame `resumo` com esses registros.
8. Se existir coluna `Estrutura`, aplica um filtro tolerante para **manter apenas linhas relacionadas a _central de inicios_** (normaliza texto removendo acentos e busca por palavras como `central` + qualquer forma de `inic`/`inicio`/`inicios`).
   - Se o filtro eliminar todas as linhas, o script avisa (print) e restaura o `resumo` original (evita perda acidental de dados).
9. Conta `Qtd_Repeticoes` por linha (para a aba `Todos_Enderecos`).
10. Salva o arquivo `clientes_enderecos_<YYYY-MM-DD>.xlsx` com as abas `Todos_Enderecos` e `Resumo`.

---

## 📝 Como personalizar
- Alterar número de top endereços: edite `TOP_N` (ex.: `TOP_N = 5`).
- Alterar comportamento do filtro de `Estrutura`:
  - A lista `possible_estrutura_cols` contém nomes que o script tenta detectar; se sua coluna tiver outro nome, adicione aqui.
  - O filtro procura por `central` E uma forma de `inic` (inic, inicio, inicios). Para mudar isso, edite o bloco onde é construída a variável `mask` antes de filtrar.
- Filtrar clientes bloqueados (remover do `resumo`): atualmente o script mantém todos os registros e só inclui a coluna `Bloqueado` no `Resumo`. Se quiser **remover** os bloqueados, adicione antes da criação do `resumo` algo como:

```python
if 'Bloqueado' in df.columns:
    df = df[df['Bloqueado'] != 'Sim']  # ou outra lógica conforme seus valores
```

- Ocultar a aba `Todos_Enderecos` na saída Excel: o script atual salva as duas abas; para esconder a aba automaticamente (requer `openpyxl`) você pode usar este trecho após escrever as abas:

```python
from openpyxl import load_workbook
wb = load_workbook(nome_arquivo)
if 'Todos_Enderecos' in wb.sheetnames:
    wb['Todos_Enderecos'].sheet_state = 'hidden'
wb.save(nome_arquivo)
```

Ou, ao usar `pd.ExcelWriter(..., engine='openpyxl')`, acessar `writer.book[...]` e ajustar `sheet_state = 'hidden'`.

---

## ✅ Saída esperada
- Arquivo: `clientes_enderecos_YYYY-MM-DD.xlsx`
- Aba `Todos_Enderecos`: seus dados originais com colunas auxiliares.
- Aba `Resumo`: linhas por cliente para os endereços top, incluindo `Bloqueado` e `Estrutura` (quando presentes). Quando a coluna `Estrutura` existir, o `Resumo` é filtrado para *central de inicios* (com tolerância a acentos e variações); se esse filtro remover todas as linhas, o script restaura o `Resumo` sem filtro e avisa.

---

## 🔎 Dicas de depuração
- Se `Resumo` estiver vazio:
  - Verifique se existem dados nos top endereços (pode haver diferenças de normalização em `RuaResidencial`).
  - Se tiver a coluna `Estrutura`, verifique se os valores contêm a expressão esperada (ex.: "Central de Inicios", "central de inícios"). O filtro é tolerante, mas você pode torná-lo mais permissivo ou desativá-lo temporariamente para teste.
- Se não é detectada a coluna de cliente ou código, verifique os cabeçalhos exatos (o script tenta normalizar nomes, mas nomes muito diferentes requerem adicionar à lista de candidatos).

---

## ℹ️ Observações finais
- Posso adaptar o script para:
  - Aceitar o caminho do arquivo como argumento de linha de comando
  - Tornar o filtro de `Estrutura` configurável via parâmetro
  - Adicionar testes automáticos ou um modo **dry-run** que gera apenas `Resumo`

Se quiser, eu já gero uma versão do `enderecos.xlsx` de exemplo e executo o script para mostrar o resultado. Deseja que eu faça isso agora? 😊
