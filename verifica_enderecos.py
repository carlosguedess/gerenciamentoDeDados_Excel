import pandas as pd
from datetime import datetime
import unicodedata

# Vai Ler o arquivo Excel (precisa ter colunas "Cliente" e "Endereço")
df = pd.read_excel("enderecos.xlsx")

# Normalizar texto dos endereços
df["RuaResidencial"] = df["RuaResidencial"].str.strip().str.lower()

# Marcar duplicados com texto em vez de True/False
df["Endereco_Repetido"] = df.duplicated(subset="RuaResidencial", keep=False).map({True: "Duplicado", False: "Único"})

# Calcular os 10 endereços que mais se repetem (top 10)
contagem = df["RuaResidencial"].value_counts()
TOP_N = 10
top_n = contagem.head(TOP_N)
top_enderecos = top_n.index.tolist()

# Criar DataFrame resumo com uma linha por cliente para cada endereço top (estilo Excel)
possible_name_cols = ["Cliente", "Nome", "nome", "nome_cliente", "Nome_Cliente"]
possible_code_cols = ["Código", "Codigo", "codigo", "ID", "Id", "id", "ClienteID", "Cliente_ID", "Codigo_Cliente", "Cod", "CodigoRevendedor", "Codigo_Revendedor", "codigo_revendedor"]

# Possíveis nomes para coluna de bloqueio (se existir)
possible_block_cols = ["Bloqueado", "Bloqueio", "bloqueado", "bloqueio", "blocked"]

# Possíveis nomes para coluna de estrutura (se existir)
possible_estrutura_cols = [
    "Estrutura", "Estrutra", "estrutura", "estrutra", "setor", "local",
    "EstruturaComercial", "Estrutura_Comercial", "estruturacomercial", "estrutura_comercial", "estrutura comercial"
]

# Detectar coluna de nome/código com fallback por normalização e substrings
name_col = next((c for c in possible_name_cols if c in df.columns), None)
code_col = next((c for c in possible_code_cols if c in df.columns), None)

if name_col is None:
    for c in df.columns:
        cn = str(c).lower().replace(' ', '').replace('_', '')
        if any(n.lower().replace(' ', '').replace('_', '') == cn for n in possible_name_cols):
            name_col = c
            break

if code_col is None:
    # tentativa por normalização direta
    for c in df.columns:
        cn = str(c).lower().replace(' ', '').replace('_', '')
        if any(k.lower().replace(' ', '').replace('_', '') == cn for k in possible_code_cols):
            code_col = c
            break
    # fallback por substring (qualquer coluna contendo 'codigo', 'cod', 'id' ou 'revendedor')
    if code_col is None:
        for c in df.columns:
            cn = str(c).lower().replace(' ', '').replace('_', '')
            if any(sub in cn for sub in ['codigo', 'cod', 'id', 'revendedor']):
                code_col = c
                break

# Detectar coluna de bloqueio (se existir)
block_col = next((c for c in possible_block_cols if c in df.columns), None)
if block_col is None:
    for c in df.columns:
        cn = str(c).lower().replace(' ', '').replace('_', '')
        if any(sub in cn for sub in ['bloquead','bloqueio','blocked']):
            block_col = c
            break

# Função para normalizar texto (remove acentos e deixa em lowercase)
def normalize_text(value):
    try:
        s = str(value)
    except Exception:
        return ""
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower().strip()

# Detectar coluna de estrutura (se existir)
estrutura_col = next((c for c in possible_estrutura_cols if c in df.columns), None)
if estrutura_col is None:
    for c in df.columns:
        cn = str(c).lower().replace(' ', '').replace('_', '')
        if 'estrut' in cn or 'setor' in cn or 'local' in cn or 'centra' in cn:
            estrutura_col = c
            break

# Nome a usar no resumo para a coluna de estrutura (preserva nome real se detectada)
estrutura_output_col = estrutura_col if estrutura_col is not None else 'Estrutura'

records = []
for endereco in top_enderecos:
    subset = df[df["RuaResidencial"] == endereco]
    if subset.empty:
        continue
    qtd = int(contagem.get(endereco, 0))
    
    # criar uma linha por cliente (preservando a ordem original)
    for _, row in subset.iterrows():
        cliente = ""
        CodigoRevendedor = ""
        bloqueado = ""
        estrutura_val = ""
        if name_col and pd.notna(row.get(name_col)):
            cliente = str(row.get(name_col))
        if code_col and pd.notna(row.get(code_col)):
            CodigoRevendedor = str(row.get(code_col))
        if block_col and pd.notna(row.get(block_col)):
            bloqueado = str(row.get(block_col))
        if estrutura_col and pd.notna(row.get(estrutura_col)):
            estrutura_val = str(row.get(estrutura_col))
        registros = {
            "Endereco": endereco,
            "Quantidade": qtd,
            "Cliente": cliente,
            "CodigoRevendedor": CodigoRevendedor,
            "Bloqueado": bloqueado,
            estrutura_output_col: estrutura_val
        }
        records.append(registros)

if records:
    resumo = pd.DataFrame(records)
else:
    # se não houver endereços, criar resumo vazio (inclui coluna de bloqueio e estrutura se houver)
    cols = ["Endereco","Quantidade","Cliente","CodigoRevendedor","Bloqueado",estrutura_output_col]
    resumo = pd.DataFrame(columns=cols)

# Se foi detectada coluna de estrutura, filtrar o resumo para manter apenas linhas com 'central de inicios'
if estrutura_output_col in resumo.columns:
    try:
        resumo['_estrutura_norm'] = resumo[estrutura_output_col].apply(lambda v: normalize_text(v) if pd.notna(v) else '')
        pre_count = len(resumo)
        # filtro mais tolerante: precisa conter 'central' E alguma forma de 'inicio' ('inic', 'inicio', 'inicios')
        mask = (
            resumo['_estrutura_norm'].str.contains('central') & (
                resumo['_estrutura_norm'].str.contains('inic') |
                resumo['_estrutura_norm'].str.contains('inicio') |
                resumo['_estrutura_norm'].str.contains('inicios')
            )
        )
        resumo = resumo[mask].copy()
        resumo = resumo.drop(columns=['_estrutura_norm'])
        post_count = len(resumo)
        if pre_count > 0 and post_count == 0:
            # Se o filtro eliminou todas as linhas, avisar e restaurar o resumo original (sem filtro)
            print(f"Aviso: filtro por 'central de inicios' retornou 0 linhas (antes: {pre_count}); mantendo resumo sem filtro.")
            resumo = pd.DataFrame(records)
    except Exception:
        # Se houver qualquer erro na filtragem, não interromper o processo
        pass

# Contar repetições por linha
df["Qtd_Repeticoes"] = df.groupby("RuaResidencial")["RuaResidencial"].transform("count")


# Nome do arquivo com data
data_hoje = datetime.now().strftime("%Y-%m-%d")
nome_arquivo = f"clientes_enderecos_{data_hoje}.xlsx"

# Salvar em duas abas: todos os dados + resumo
with pd.ExcelWriter(nome_arquivo) as writer:
    df.to_excel(writer, sheet_name="Todos_Enderecos", index=False)
    resumo.to_excel(writer, sheet_name="Resumo", index=False)
   
print("Processo finalizado com sucesso!")
