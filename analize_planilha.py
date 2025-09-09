import pandas as pd
import matplotlib.pyplot as plt
import re


# Carrega o arquivo Excel
xl_file = pd.ExcelFile('CONTROLE ESTACAO.xlsx')


# === Funções utilitárias ===
def get_col_indices(header_row, nomes):
    indices = {nome: None for nome in nomes}
    for idx, col_name in enumerate(header_row):
        nome_col = str(col_name).strip().upper()
        if nome_col in indices:
            indices[nome_col] = idx
    return indices

# Lê todas as abas da planilha
dfs = {}
abas_com_erro = []
for i, nome_aba in enumerate(xl_file.sheet_names):
    try:
        df = xl_file.parse(i)
        dfs[nome_aba] = df
    except Exception as e:
        print(f"[AVISO] Erro ao ler a aba de índice {i} ('{nome_aba}'): {e}")
        abas_com_erro.append((i, nome_aba))
if abas_com_erro:
    print(f"\nAs seguintes abas não puderam ser lidas e foram ignoradas (para corrigir, desligue os filtros):")
    for idx, nome in abas_com_erro:
        print(f"  Índice: {idx} | Nome: {nome}")

# === Geração da blacklist de históricos inválidos ===
blacklist_historicos = set()
for nome_aba, df_full in dfs.items():
    header_row = df_full.iloc[0]
    indices = get_col_indices(header_row, ['HISTÓRICO'])
    historico_idx = indices['HISTÓRICO']
    if historico_idx is None: continue
    df = df_full.iloc[1:].reset_index(drop=True)
    historicos = df.iloc[:, historico_idx].astype(str)
    for h in historicos:
        if 'PROT' not in h.upper() and 'IA-' not in h.upper():
            blacklist_historicos.add(h.strip())

# === Análise por Animal (Vaca) ===
ultima_aba = xl_file.sheet_names[-1]
df_full = dfs[ultima_aba]
header_row = df_full.iloc[0]
indices = get_col_indices(header_row, ['ANIMAL'])
animal_idx = indices['ANIMAL']
df = df_full.iloc[1:].reset_index(drop=True)
animal_ids = df.iloc[:, animal_idx].astype(str).str.strip().unique().tolist()
animal_ids = [aid for aid in animal_ids if aid]

vacas_data = []
num_animais_ignorados = 0
prot_labels = [f"{i}PROT" for i in range(1, 32)]

for animal_id in animal_ids:
    animal_invalido = False
    for nome_aba, df_full_check in dfs.items():
        header_row_check = df_full_check.iloc[0]
        indices_check = get_col_indices(header_row_check, ['ANIMAL', 'HISTÓRICO'])
        animal_idx_check, historico_idx_check = indices_check['ANIMAL'], indices_check['HISTÓRICO']
        if animal_idx_check is None or historico_idx_check is None: continue
        df_check = df_full_check.iloc[1:]
        mask = df_check.iloc[:, animal_idx_check].astype(str).str.strip() == animal_id
        historicos = df_check[mask].iloc[:, historico_idx_check].astype(str)
        for h in historicos:
            if h.strip() in blacklist_historicos:
                animal_invalido = True
                break
        if animal_invalido: break
    
    if animal_invalido:
        num_animais_ignorados += 1
        continue

    pesos_machos, pesos_femeas = [], []
    qtd_machos, qtd_femeas = 0, 0
    total_tentativas, total_prenhezes, total_abortos = 0, 0, 0
    
    prot_total_vaca = {prot: 0 for prot in prot_labels}
    prot_prenhez_vaca = {prot: 0 for prot in prot_labels}

    for nome_aba, df_full_data in dfs.items():
        header_row_data = df_full_data.iloc[0]
        indices_data = get_col_indices(header_row_data, ['HISTÓRICO', 'SITUAÇÃO', 'PESO 205', 'ANIMAL', 'SEXO'])
        historico_idx, situacao_idx, peso_idx, animal_idx, sexo_idx = (
            indices_data['HISTÓRICO'], indices_data['SITUAÇÃO'], indices_data['PESO 205'], indices_data['ANIMAL'], indices_data['SEXO']
        )
        if animal_idx is None: continue
        
        df_data = df_full_data.iloc[1:]
        mask_animal = df_data.iloc[:, animal_idx].astype(str).str.strip() == animal_id
        df_animal = df_data[mask_animal]
        if df_animal.empty: continue

        if historico_idx is not None:
            valid_mask = ~df_animal.iloc[:, historico_idx].astype(str).str.strip().isin(blacklist_historicos)
            df_animal = df_animal[valid_mask]
            if df_animal.empty: continue
        
            historicos_vaca = df_animal.iloc[:, historico_idx].astype(str).str.replace(' ', '', regex=False).str.upper()
            for prot in prot_labels:
                prot_total_vaca[prot] += historicos_vaca.str.contains(rf'(?<!\d){prot}(?!\d)', regex=True, na=False).sum()
                prot_prenhez_vaca[prot] += historicos_vaca.str.contains(rf'(?<!\d){prot}-(?:P|AB)', regex=True, na=False).sum()

        if peso_idx is not None and sexo_idx is not None:
            pesos = pd.to_numeric(df_animal.iloc[:, peso_idx], errors='coerce')
            sexos = df_animal.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
            pesos_machos.extend(pesos[sexos == 'M'].dropna().tolist())
            pesos_femeas.extend(pesos[sexos == 'F'].dropna().tolist())
            qtd_machos += (sexos == 'M').sum()
            qtd_femeas += (sexos == 'F').sum()

        if situacao_idx is not None:
            situacoes = df_animal.iloc[:, situacao_idx].astype(str)
            total_tentativas += len(situacoes)
            total_prenhezes += situacoes.isin(['P', 'AB', 'P2', 'REAB']).sum()
            total_abortos += situacoes.isin(['AB', 'REAB']).sum()
    
    peso_medio_global = (sum(pesos_machos) + sum(pesos_femeas)) / (len(pesos_machos) + len(pesos_femeas)) if (pesos_machos or pesos_femeas) else None
    
    vaca_info = {
        'animal_id': animal_id, 'total_tentativas': total_tentativas, 'total_prenhezes': total_prenhezes,
        'total_abortos': total_abortos, 'taxa_global': total_prenhezes / total_tentativas if total_tentativas > 0 else None,
        'peso_medio_global': peso_medio_global, 'qtd_machos': qtd_machos, 'qtd_femeas': qtd_femeas,
    }

    for prot in prot_labels:
        total = prot_total_vaca[prot]
        prenhez = prot_prenhez_vaca[prot]
        vaca_info[f'{prot}_total_vaca'] = total
        vaca_info[f'{prot}_prenhez_vaca'] = prenhez
        vaca_info[f'{prot}_taxa_vaca'] = prenhez / total if total > 0 else None

    vacas_data.append(vaca_info)

# --- Coleta estatísticas por estação (ANO) ---
estacoes_data = []
for nome_aba, df_full in dfs.items():
    header_row = df_full.iloc[0]
    indices = get_col_indices(header_row, ['ANIMAL', 'PESO 205', 'SEXO', 'SITUAÇÃO', 'HISTÓRICO', 'CATEGORIA'])
    animal_idx, peso_idx, sexo_idx, situacao_idx, historico_idx, categoria_idx = (
        indices['ANIMAL'], indices['PESO 205'], indices['SEXO'], indices['SITUAÇÃO'], indices['HISTÓRICO'], indices['CATEGORIA']
    )
    df = df_full.iloc[1:].reset_index(drop=True)
    if historico_idx is not None:
        historicos_aba = df.iloc[:, historico_idx].astype(str).str.strip()
        mask_validos = ~historicos_aba.isin(blacklist_historicos)
        df = df[mask_validos]
    num_registros = len(df)
    situacoes = df.iloc[:, situacao_idx].astype(str) if situacao_idx is not None else pd.Series([])
    total_concepcoes = situacoes.isin(['P', 'AB', 'P2', 'REAB']).sum()
    taxa_prenhez_geral = total_concepcoes / len(situacoes) if len(situacoes) > 0 else None

    prot_stats = {}
    if historico_idx is not None:
        historicos = df.iloc[:, historico_idx].astype(str).str.replace(' ', '', regex=False).str.upper()
        for prot in prot_labels:
            mask_total = historicos.str.contains(rf'(?<!\d){prot}(?!\d)', regex=True, na=False)
            total_prot = mask_total.sum()
            if total_prot > 0:
                mask_prenhez = historicos.str.contains(rf'(?<!\d){prot}-(?:P|AB)', regex=True, na=False)
                
                # ########################################################################## #
                # ### LÓGICA DE ABORTO CORRIGIDA ###
                # Agora, um aborto é contado se o protocolo estiver na linha E "-AB"
                # também estiver na linha, cobrindo casos como "4PROT-P-AB".
                mask_contem_aborto_geral = historicos.str.contains(r'-AB', na=False)
                mask_aborto_final = mask_total & mask_contem_aborto_geral
                abortos_prot = mask_aborto_final.sum()
                # ### FIM DA CORREÇÃO ###
                # ########################################################################## #

                prenhezes_prot = mask_prenhez.sum()
                taxa_prot = prenhezes_prot / total_prot if total_prot > 0 else 0
                
                prot_stats[prot] = {
                    'total': total_prot, 'prenhezes': prenhezes_prot,
                    'taxa': taxa_prot, 'abortos': abortos_prot
                }
    
    categoria_stats = {}
    if categoria_idx is not None and situacao_idx is not None:
        categorias = df.iloc[:, categoria_idx].dropna().astype(str).str.strip().unique()
        for cat in categorias:
            if not cat: continue
            mask_cat = df.iloc[:, categoria_idx].astype(str).str.strip() == cat
            situacoes_cat = df.loc[mask_cat, df.columns[situacao_idx]].astype(str)
            total_cat = len(situacoes_cat)
            prenhezes_cat = situacoes_cat.isin(['P', 'AB', 'P2', 'REAB']).sum()
            taxa_cat = prenhezes_cat / total_cat if total_cat > 0 else None
            categoria_stats[cat] = {'total': total_cat, 'prenhezes': prenhezes_cat, 'taxa': taxa_cat}

    estacoes_data.append({
        'estacao': nome_aba, 'total_registros': num_registros, 'total_concepcoes': total_concepcoes,
        'total_abortos': situacoes.isin(['AB', 'REAB']).sum(),
        'peso_medio': pd.to_numeric(df.iloc[:, peso_idx], errors='coerce').mean() if peso_idx is not None else None,
        'qtd_machos': (df.iloc[:, sexo_idx].astype(str).str.upper().str.strip() == 'M').sum() if sexo_idx is not None else 0,
        'qtd_femeas': (df.iloc[:, sexo_idx].astype(str).str.upper().str.strip() == 'F').sum() if sexo_idx is not None else 0,
        'taxa_prenhez_geral': taxa_prenhez_geral, 'prots': prot_stats, 'categorias': categoria_stats
    })

def expand_stats(data, prot_labels, categoria_labels):
    expanded = []
    for row in data:
        new_row = row.copy()
        if 'prots' in row:
            for prot in prot_labels:
                stats = row['prots'].get(prot, {})
                new_row[f'{prot}_total'] = stats.get('total')
                new_row[f'{prot}_prenhezes'] = stats.get('prenhezes')
                new_row[f'{prot}_taxa'] = stats.get('taxa')
                new_row[f'{prot}_abortos'] = stats.get('abortos')
            del new_row['prots']
        if 'categorias' in row:
            for cat in categoria_labels:
                stats = row['categorias'].get(cat, {})
                new_row[f'{cat}_total'] = stats.get('total')
                new_row[f'{cat}_prenhezes'] = stats.get('prenhezes')
                new_row[f'{cat}_taxa'] = stats.get('taxa')
            del new_row['categorias']
        expanded.append(new_row)
    return expanded

categoria_labels = set()
for est in estacoes_data:
    if 'categorias' in est:
        categoria_labels.update(est['categorias'].keys())
categoria_labels = sorted(list(categoria_labels))

estacoes_data_expanded = expand_stats(estacoes_data, prot_labels, categoria_labels)

with pd.ExcelWriter('saida_analise.xlsx') as writer:
    pd.DataFrame(vacas_data).to_excel(writer, sheet_name='Vacas', index=False)
    pd.DataFrame(estacoes_data_expanded).to_excel(writer, sheet_name='Estacoes', index=False)

print("\nArquivo 'saida_analise.xlsx' gerado com sucesso!")
print("A lógica de contagem de abortos por protocolo foi ajustada para máxima precisão.")

print("\nHistóricos inválidos encontrados (ignorados nas contagens):")
if blacklist_historicos:
    for h in sorted(list(blacklist_historicos)):
        print(f"  - {h}")
else:
    print("  Nenhum histórico inválido encontrado.")
print(f"\nTotal de vacas ignoradas na análise individual (por terem histórico inválido em algum momento): {num_animais_ignorados}")
print("\n--- Resumo por Estação ---")
for estacao in estacoes_data:
    nome = estacao['estacao']
    registros = estacao['total_registros'] 
    taxa = estacao['taxa_prenhez_geral']
    if taxa is not None:
        print(f"Estação: {nome:<10} | Total de Registros: {registros:<5} | Taxa de Prenhez Geral: {taxa:.2%}")
    else:
        print(f"Estação: {nome:<10} | Total de Registros: {registros:<5} | Taxa de Prenhez Geral: N/A")