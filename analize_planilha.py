# --- Bloco de Importação de Bibliotecas ---
import pandas as pd
import matplotlib.pyplot as plt
import os
from tqdm import tqdm
import re # Importado para usar expressões regulares na contagem de protocolos

# --- Carregamento do Arquivo Excel ---
try:
    xl_file = pd.ExcelFile('CONTROLE ESTAÇÃO.ATUAL.xlsx') 
except FileNotFoundError:
    print("ERRO: O arquivo 'CONTROLE ESTAÇÃO.ATUAL.xlsx' não foi encontrado. Verifique o nome e o local.")
    exit()


# --- Seção de Interação com o Usuário ---
abas_disponiveis = xl_file.sheet_names
print(f"Abas encontradas na planilha: {abas_disponiveis}")
num_abas_a_usar = input(f"\nDigite o número de abas mais recentes a serem analisadas (ex: 2 para usar as duas últimas). \nPressione Enter para analisar todas as {len(abas_disponiveis)} abas: ")
abas_selecionadas = []
try:
    if num_abas_a_usar.strip() == "":
        abas_selecionadas = abas_disponiveis
        print(f"Analisando todas as {len(abas_selecionadas)} abas.")
    else:
        num = int(num_abas_a_usar)
        if num > 0:
            abas_selecionadas = abas_disponiveis[-num:]
            print(f"Analisando as últimas {len(abas_selecionadas)} abas: {abas_selecionadas}")
        else:
            abas_selecionadas = abas_disponiveis
            print("Número inválido. Analisando todas as abas por padrão.")
except ValueError:
    abas_selecionadas = abas_disponiveis
    print("Entrada inválida. Analisando todas as abas por padrão.")

# --- Funções Utilitárias ---
def get_col_indices(header_row, nomes):
    indices = {}
    nomes_a_buscar = {str(nome).strip().upper() for nome in nomes}
    nomes_encontrados_map = {} 

    for idx, col_name in enumerate(header_row):
        nome_col_header = str(col_name).strip().upper()
        if nome_col_header in nomes_a_buscar:
            for nome_original in nomes:
                if str(nome_original).strip().upper() == nome_col_header:
                    indices[nome_original] = idx
                    nomes_encontrados_map[nome_original] = nome_col_header
                    break
    
    for nome_original in nomes:
        if nome_original not in indices:
             indices[nome_original] = None
             
    return indices

# --- Leitura e Carregamento dos Dados ---
dfs = {}
abas_com_erro = []
print("\nCarregando dados das abas selecionadas...")
for nome_aba in tqdm(abas_selecionadas, desc="Carregando Abas"):
    try:
        df = xl_file.parse(nome_aba, dtype=str) 
        dfs[nome_aba] = df
    except Exception as e:
        print(f"[AVISO] Erro ao ler a aba '{nome_aba}': {e}")
        abas_com_erro.append(nome_aba)
if abas_com_erro:
    print(f"\nAs seguintes abas não puderam ser lidas e foram ignoradas: {abas_com_erro}")

# --- Geração da Blacklist de Históricos ---
blacklist_historicos = set()
for nome_aba, df_full in dfs.items():
    try:
        if df_full.empty: continue 
        # Garante que header_row seja uma Series antes de usar .astype(str)
        if isinstance(df_full.iloc[0], pd.Series):
             header_row = df_full.iloc[0].astype(str)
        else: # Se for DataFrame (caso raro), pega a primeira linha como lista
             header_row = df_full.iloc[0].values.astype(str)

        indices = get_col_indices(header_row, ['HISTÓRICO'])
        historico_idx = indices.get('HISTÓRICO')
        if historico_idx is None: continue
        
        df_data = df_full.iloc[1:].reset_index(drop=True)
        if not df_data.empty:
            if historico_idx < len(df_data.columns):
                 historicos = df_data.iloc[:, historico_idx].astype(str)
                 for h in historicos:
                      h_upper = h.upper()
                      if not ('PROT' in h_upper or 'IA-' in h_upper):
                           blacklist_historicos.add(h.strip())
            #else:
                 #print(f"[AVISO] Índice 'HISTÓRICO' ({historico_idx}) fora dos limites na aba '{nome_aba}'.") # Descomente se precisar debugar

    except IndexError:
         print(f"[AVISO] Aba '{nome_aba}' parece vazia ou só cabeçalho. Ignorando para blacklist.")
    except Exception as e:
        print(f"[AVISO] Erro processando blacklist na aba '{nome_aba}': {e}")


# --- Análise por Animal (Vaca) ---
# (Esta seção permanece inalterada, focada nos dados gerais da vaca)
animal_ids = []
if abas_selecionadas:
    try:
        ultima_aba = abas_selecionadas[-1]
        if ultima_aba in dfs and not dfs[ultima_aba].empty:
            df_full_last = dfs[ultima_aba]
            # Garante que header_row seja uma Series
            if isinstance(df_full_last.iloc[0], pd.Series):
                 header_row_last = df_full_last.iloc[0].astype(str)
            else:
                 header_row_last = df_full_last.iloc[0].values.astype(str)

            indices_last = get_col_indices(header_row_last, ['ANIMAL'])
            animal_idx_last = indices_last.get('ANIMAL')
            if animal_idx_last is not None:
                df_data_last = df_full_last.iloc[1:].reset_index(drop=True)
                if not df_data_last.empty and animal_idx_last < len(df_data_last.columns):
                    animal_ids = df_data_last.iloc[:, animal_idx_last].dropna().astype(str).str.strip().unique().tolist()
                    animal_ids = [aid for aid in animal_ids if aid]
    except IndexError:
         print(f"[AVISO] Última aba '{ultima_aba}' vazia ou só cabeçalho.")
    except Exception as e:
        print(f"[ERRO] Carregando lista de animais da última aba '{ultima_aba}': {e}")

vacas_data = []
num_animais_ignorados = 0
prot_labels = [f"{i}PROT" for i in range(1, 31)]

print("\nIniciando análise por animal...")
for animal_id in tqdm(animal_ids, desc="Analisando Vacas"):
    animal_invalido = False
    try:
        for nome_aba, df_full_check in dfs.items():
             if df_full_check.empty: continue
             # Garante header como Series
             if isinstance(df_full_check.iloc[0], pd.Series):
                  header_row_check = df_full_check.iloc[0].astype(str)
             else:
                  header_row_check = df_full_check.iloc[0].values.astype(str)
                  
             indices_check = get_col_indices(header_row_check, ['ANIMAL', 'HISTÓRICO'])
             animal_idx_check = indices_check.get('ANIMAL')
             historico_idx_check = indices_check.get('HISTÓRICO')
             if animal_idx_check is None or historico_idx_check is None: continue

             df_check_data = df_full_check.iloc[1:]
             if df_check_data.empty: continue
             if animal_idx_check >= len(df_check_data.columns) or historico_idx_check >= len(df_check_data.columns): continue

             df_animal_check = df_check_data[df_check_data.iloc[:, animal_idx_check].astype(str).str.strip() == animal_id]
             if not df_animal_check.empty:
                  if df_animal_check.iloc[:, historico_idx_check].astype(str).str.strip().isin(blacklist_historicos).any():
                       animal_invalido = True
                       break 
        
        if animal_invalido:
            num_animais_ignorados += 1
            continue
            
    except Exception as e:
         print(f"[AVISO] Erro na checagem de blacklist para {animal_id}: {e}.")
         continue 

    pesos_machos, pesos_femeas = [], []
    qtd_machos, qtd_femeas = 0, 0
    numero_de_estacoes, total_prenhezes, total_abortos = 0, 0, 0
    ultima_categoria = "N/A"

    for nome_aba, df_full_data in dfs.items():
        try:
            if df_full_data.empty: continue
            # Garante header como Series
            if isinstance(df_full_data.iloc[0], pd.Series):
                 header_row_data = df_full_data.iloc[0].astype(str)
            else:
                 header_row_data = df_full_data.iloc[0].values.astype(str)

            indices_data = get_col_indices(header_row_data, ['HISTÓRICO', 'SITUAÇÃO', 'PESO 205', 'ANIMAL', 'SEXO', 'CATEGORIA'])
            historico_idx = indices_data.get('HISTÓRICO')
            situacao_idx = indices_data.get('SITUAÇÃO')
            peso_idx = indices_data.get('PESO 205')
            animal_idx = indices_data.get('ANIMAL')
            sexo_idx = indices_data.get('SEXO')
            categoria_idx = indices_data.get('CATEGORIA')

            if animal_idx is None: continue
            
            df_data = df_full_data.iloc[1:]
            if df_data.empty or animal_idx >= len(df_data.columns): continue

            mask_animal = df_data.iloc[:, animal_idx].astype(str).str.strip() == animal_id
            df_animal = df_data[mask_animal]
            if df_animal.empty: continue

            if categoria_idx is not None and categoria_idx < len(df_animal.columns):
                categorias_encontradas = df_animal.iloc[:, categoria_idx].dropna().astype(str).str.strip().unique()
                if len(categorias_encontradas) > 0:
                    ultima_categoria = categorias_encontradas[-1]

            df_animal_valid = df_animal # Assume válido inicialmente
            if historico_idx is not None and historico_idx < len(df_animal.columns):
                valid_mask = ~df_animal.iloc[:, historico_idx].astype(str).str.strip().isin(blacklist_historicos)
                df_animal_valid = df_animal[valid_mask]
                if df_animal_valid.empty: continue 
            
            if peso_idx is not None and sexo_idx is not None and \
               peso_idx < len(df_animal_valid.columns) and sexo_idx < len(df_animal_valid.columns):
                pesos = pd.to_numeric(df_animal_valid.iloc[:, peso_idx], errors='coerce')
                sexos = df_animal_valid.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
                pesos_machos.extend(pesos[sexos == 'M'].dropna().tolist())
                pesos_femeas.extend(pesos[sexos == 'F'].dropna().tolist())
                qtd_machos += (sexos == 'M').sum()
                qtd_femeas += (sexos == 'F').sum()

            if situacao_idx is not None and situacao_idx < len(df_animal_valid.columns):
                situacoes = df_animal_valid.iloc[:, situacao_idx].astype(str)
                numero_de_estacoes += len(situacoes)
                total_prenhezes += situacoes.isin(['P', 'AB', 'P2', 'REAB']).sum()
                total_abortos += situacoes.isin(['AB', 'REAB']).sum()
        except Exception as e:
            print(f"[AVISO] Erro processando dados do animal {animal_id} na aba '{nome_aba}': {e}")

    
    peso_medio_desmame = (sum(pesos_machos) + sum(pesos_femeas)) / (len(pesos_machos) + len(pesos_femeas)) if (pesos_machos or pesos_femeas) else None
    concepcao_por_estação_num = total_prenhezes / numero_de_estacoes if numero_de_estacoes > 0 else None
    
    vaca_info = {
        'animal_id': animal_id, 'ultima_categoria': ultima_categoria,
        'numero_de_estacoes': numero_de_estacoes, 'total_prenhezes': total_prenhezes,
        'total_abortos': total_abortos, 
        'concepcao_por_estação': concepcao_por_estação_num, 
        'peso_medio_desmame': peso_medio_desmame, 
        'qtd_machos': qtd_machos, 'qtd_femeas': qtd_femeas,
    }
    vacas_data.append(vaca_info)

# --- Coleta de Dados para as Múltiplas Abas de Estação ---
estacoes_global_data = []
analise_categoria_data = []
estacoes_protocolo_data = []

print("\nIniciando análise por estação...")
for nome_aba, df_full in tqdm(dfs.items(), desc="Analisando Estações"):
    try:
        if df_full.empty: continue
        # Garante header como Series
        if isinstance(df_full.iloc[0], pd.Series):
             header_row = df_full.iloc[0].astype(str).str.upper()
        else:
             header_row = df_full.iloc[0].values.astype(str).str.upper()

        indices = get_col_indices(header_row, ['ANIMAL', 'PESO 205', 'SEXO', 'SITUAÇÃO', 'HISTÓRICO', 'CATEGORIA', 'DATA IA'])
        animal_idx = indices.get('ANIMAL')
        peso_idx = indices.get('PESO 205')
        sexo_idx = indices.get('SEXO')
        situacao_idx = indices.get('SITUAÇÃO')
        historico_idx = indices.get('HISTÓRICO')
        categoria_idx = indices.get('CATEGORIA')
        data_ia_idx = indices.get('DATA IA')
        
        if historico_idx is None or situacao_idx is None or categoria_idx is None:
            print(f"[AVISO] Aba '{nome_aba}' ignorada: Faltam colunas essenciais.")
            continue
            
        df_data = df_full.iloc[1:].reset_index(drop=True)
        if df_data.empty: continue
        if historico_idx >= len(df_data.columns) or situacao_idx >= len(df_data.columns) or categoria_idx >= len(df_data.columns):
            print(f"[AVISO] Aba '{nome_aba}' ignorada: Índices de colunas essenciais fora dos limites.")
            continue
            
        df_valid = df_data[~df_data.iloc[:, historico_idx].astype(str).str.strip().isin(blacklist_historicos)]
        if df_valid.empty: continue

        # --- Cálculos Globais da Estação ---
        situacoes_global = df_valid.iloc[:, situacao_idx].astype(str).str.upper()
        historicos_estacao = df_valid.iloc[:, historico_idx].astype(str).str.replace(' ', '', regex=False).str.upper()
        
        taxa_prenhez_geral_num = situacoes_global.str.upper().isin(['P', 'AB', 'P2', 'REAB']).sum() / len(df_valid) if len(df_valid) > 0 else None
        
        peso_medio_machos_global, peso_medio_femeas_global = None, None
        data_media_ia_geral, data_media_ia_machos, data_media_ia_femeas = None, None, None
        # ... (cálculos de peso e data global - sem alterações) ...
        if peso_idx is not None and sexo_idx is not None and peso_idx < len(df_valid.columns) and sexo_idx < len(df_valid.columns):
            pesos = pd.to_numeric(df_valid.iloc[:, peso_idx], errors='coerce')
            sexos = df_valid.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
            peso_medio_machos_global = pesos[sexos == 'M'].mean()
            peso_medio_femeas_global = pesos[sexos == 'F'].mean()
        if data_ia_idx is not None and data_ia_idx < len(df_valid.columns):
             datas_ia = pd.to_datetime(df_valid.iloc[:, data_ia_idx], errors='coerce')
             data_media_geral = datas_ia.dropna().mean()
             if sexo_idx is not None and sexo_idx < len(df_valid.columns):
                 sexos_datas = df_valid.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
                 data_media_ia_machos = datas_ia[sexos_datas == 'M'].dropna().mean()
                 data_media_ia_femeas = datas_ia[sexos_datas == 'F'].dropna().mean()


        # --- Novos Cálculos: Totais Brutos de Protocolo (Global) ---
        total_protocolos_aplicados_global = historicos_estacao.str.upper().str.count(r'\d+PROT').sum() 
        total_protocolos_sucesso_global = historicos_estacao.str.upper().str.contains(r'PROT-(?:P|AB|P2)').sum()
        # --- Fim dos Novos Cálculos ---

        estacoes_global_data.append({
            'estacao': nome_aba, 'total_registros': len(df_valid),
            'total_concepcoes': situacoes_global.isin(['P', 'AB', 'P2', 'REAB']).sum(),
            'total_abortos': situacoes_global.isin(['AB', 'REAB']).sum(),
            'taxa_prenhez_geral': taxa_prenhez_geral_num,
            'total_protocolos_aplicados': total_protocolos_aplicados_global, # <-- NOVA COLUNA
            'total_protocolos_sucesso': total_protocolos_sucesso_global,   # <-- NOVA COLUNA
            # 'taxa_geral_protocolos' já estava sendo calculada implicitamente, vamos manter
            'taxa_geral_protocolos': (total_protocolos_sucesso_global / total_protocolos_aplicados_global) if total_protocolos_aplicados_global > 0 else None,
            'peso_medio_machos_geral': peso_medio_machos_global,
            'peso_medio_femeas_geral': peso_medio_femeas_global,
            'data_media_ia_geral': data_media_geral.strftime('%d-%m-%Y') if pd.notna(data_media_geral) else None,
            'data_media_ia_machos': data_media_ia_machos.strftime('%d-%m-%Y') if pd.notna(data_media_ia_machos) else None,
            'data_media_ia_femeas': data_media_ia_femeas.strftime('%d-%m-%Y') if pd.notna(data_media_ia_femeas) else None,
        })
        
        # --- Cálculo para "Estacoes_por_Protocolo" (Global da Estação) ---
        # (Esta parte permanece a mesma, calculando individualmente por PROT)
        prot_stats_estacao = {}
        # Recalcular totais aqui para evitar dependência do bloco anterior (embora redundante)
        total_prot_prenhez_global_indiv = 0
        total_prot_participacao_global_indiv = 0
        for prot in prot_labels:
            mask_total = historicos_estacao.str.contains(rf'(?<!\d){prot}(?!\d)', regex=True, na=False)
            total_prot = mask_total.sum()
            if total_prot > 0:
                mask_prenhez = historicos_estacao.str.contains(rf'(?<!\d){prot}-(?:P|AB)', regex=True, na=False)
                prenhez_prot = mask_prenhez.sum()
                mask_contem_aborto = historicos_estacao.str.contains(r'-AB', na=False)
                mask_aborto_final = mask_total & mask_contem_aborto
                aborto_prot = mask_aborto_final.sum()
                taxa_prot_num = prenhez_prot / total_prot
                
                prot_stats_estacao[prot] = {
                    'total': total_prot, 'prenhezes': prenhez_prot,
                    'taxa': taxa_prot_num, 'abortos': aborto_prot
                }
                total_prot_prenhez_global_indiv += prenhez_prot
                total_prot_participacao_global_indiv += total_prot
        estacoes_protocolo_data.append({'estacao': nome_aba, 'prots': prot_stats_estacao})
        
        # --- Cálculo para "Estacoes_por_Categoria" ---
        categorias = df_valid.iloc[:, categoria_idx].dropna().astype(str).str.strip().unique()
        for cat in categorias:
            if not cat: continue
            df_cat = df_valid[df_valid.iloc[:, categoria_idx].astype(str).str.strip() == cat]
            if df_cat.empty: continue 
            
            situacoes_cat = df_cat.iloc[:, situacao_idx].astype(str)
            historicos_cat = df_cat.iloc[:, historico_idx].astype(str).str.replace(' ', '', regex=False).str.upper()
            
            taxa_prenhez_cat_num = situacoes_cat.isin(['P', 'AB', 'P2', 'REAB']).sum() / len(df_cat) if len(df_cat) > 0 else None
            
            peso_medio_machos_cat, peso_medio_femeas_cat = None, None
            data_media_ia_geral_cat, data_media_ia_machos_cat, data_media_ia_femeas_cat = None, None, None
            # ... (cálculos de peso e data por categoria - sem alterações) ...
            if peso_idx is not None and sexo_idx is not None and peso_idx < len(df_cat.columns) and sexo_idx < len(df_cat.columns):
                pesos_cat = pd.to_numeric(df_cat.iloc[:, peso_idx], errors='coerce')
                sexos_cat = df_cat.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
                peso_medio_machos_cat = pesos_cat[sexos_cat == 'M'].mean()
                peso_medio_femeas_cat = pesos_cat[sexos_cat == 'F'].mean()
            if data_ia_idx is not None and data_ia_idx < len(df_cat.columns):
                 datas_ia_cat = pd.to_datetime(df_cat.iloc[:, data_ia_idx], errors='coerce')
                 data_media_geral_cat = datas_ia_cat.dropna().mean()
                 if sexo_idx is not None and sexo_idx < len(df_cat.columns):
                     sexos_datas_cat = df_cat.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
                     data_media_ia_machos_cat = datas_ia_cat[sexos_datas_cat == 'M'].dropna().mean()
                     data_media_ia_femeas_cat = datas_ia_cat[sexos_datas_cat == 'F'].dropna().mean()

            # --- Novos Cálculos: Totais Brutos de Protocolo (por Categoria) ---
            total_protocolos_aplicados_cat = historicos_cat.str.upper().str.count(r'\d+PROT').sum() 
            total_protocolos_sucesso_cat = historicos_cat.str.upper().str.contains(r'PROT-(?:P|AB|P2)').sum()
            taxa_geral_prot_cat_num = total_protocolos_sucesso_cat / total_protocolos_aplicados_cat if total_protocolos_aplicados_cat > 0 else None
            # --- Fim dos Novos Cálculos ---

            # Cálculo de Protocolo Individual por Categoria (permanece)
            prot_stats_categoria = {}
            for prot in prot_labels:
                mask_total_cat = historicos_cat.str.contains(rf'(?<!\d){prot}(?!\d)', regex=True, na=False)
                total_prot_cat = mask_total_cat.sum()
                if total_prot_cat > 0:
                    mask_prenhez_cat = historicos_cat.str.contains(rf'(?<!\d){prot}-(?:P|AB)', regex=True, na=False)
                    prenhez_prot_cat = mask_prenhez_cat.sum()
                    mask_contem_aborto_cat = historicos_cat.str.contains(r'-AB', na=False)
                    mask_aborto_final_cat = mask_total_cat & mask_contem_aborto_cat
                    aborto_prot_cat = mask_aborto_final_cat.sum()
                    taxa_prot_num_cat = prenhez_prot_cat / total_prot_cat
                    
                    prot_stats_categoria[prot] = {
                        'total': total_prot_cat, 'prenhezes': prenhez_prot_cat,
                        'taxa': taxa_prot_num_cat, 'abortos': aborto_prot_cat
                    }
            
            analise_categoria_data.append({
                'estacao': nome_aba, 'categoria': cat, 'total_registros_cat': len(df_cat),
                'total_concepcoes_cat': situacoes_cat.isin(['P', 'AB', 'P2', 'REAB']).sum(),
                'total_abortos_cat': situacoes_cat.isin(['AB', 'REAB']).sum(),
                'taxa_prenhez_cat': taxa_prenhez_cat_num, 
                'total_protocolos_aplicados_cat': total_protocolos_aplicados_cat,
                'total_protocolos_sucesso_cat': total_protocolos_sucesso_cat,
                'taxa_geral_protocolos_cat': taxa_geral_prot_cat_num,
                'peso_medio_machos_cat': peso_medio_machos_cat,
                'peso_medio_femeas_cat': peso_medio_femeas_cat,
                'data_media_ia_geral_cat': data_media_geral_cat.strftime('%d-%m-%Y') if pd.notna(data_media_geral_cat) else None,
                'data_media_ia_machos_cat': data_media_ia_machos_cat.strftime('%d-%m-%Y') if pd.notna(data_media_ia_machos_cat) else None,
                'data_media_ia_femeas_cat': data_media_ia_femeas_cat.strftime('%d-%m-%Y') if pd.notna(data_media_ia_femeas_cat) else None,
                'prots': prot_stats_categoria 
            })
    except Exception as e:
        print(f"[ERRO] Erro fatal ao processar a aba '{nome_aba}': {e}")


# --- Geração dos DataFrames Finais ---
def expand_stats_detalhado(data, prot_labels):
    expanded = []
    for row in data:
        new_row = row.copy()
        # Garante que 'prots' exista antes de tentar acessá-lo
        if 'prots' in row and isinstance(row['prots'], dict):
            for prot in prot_labels:
                # Usa .get() com um dicionário default para evitar erros se 'prots' não tiver a chave
                stats = row['prots'].get(prot, {'total': 0, 'prenhezes': 0, 'taxa': None, 'abortos': 0}) 
                new_row[f'{prot}_total'] = stats.get('total', 0) # Default 0 se chave não existir
                new_row[f'{prot}_prenhezes'] = stats.get('prenhezes', 0)
                new_row[f'{prot}_taxa'] = stats.get('taxa') # Mantém None se não existir
                new_row[f'{prot}_abortos'] = stats.get('abortos', 0)
            del new_row['prots']
        else:
             # Se 'prots' não existe ou não é dict, preenche com defaults para todas as colunas de prot
             for prot in prot_labels:
                  new_row[f'{prot}_total'] = 0
                  new_row[f'{prot}_prenhezes'] = 0
                  new_row[f'{prot}_taxa'] = None
                  new_row[f'{prot}_abortos'] = 0

        expanded.append(new_row)
    return expanded


df_vacas = pd.DataFrame(vacas_data)
if not df_vacas.empty:
    cols_primeiras = ['animal_id', 'ultima_categoria']
    cols_restantes = [col for col in df_vacas.columns if col not in cols_primeiras]
    df_vacas = df_vacas[cols_primeiras + cols_restantes]

df_estacoes_global = pd.DataFrame(estacoes_global_data)
df_estacoes_categoria = pd.DataFrame(expand_stats_detalhado(analise_categoria_data, prot_labels))
df_estacoes_protocolo = pd.DataFrame(expand_stats_detalhado(estacoes_protocolo_data, prot_labels))

# --- Geração do Arquivo Excel com Formatação ---
try:
    with pd.ExcelWriter('saida_analise.xlsx', engine='xlsxwriter') as writer:
        # Escreve os DataFrames nas abas
        if not df_vacas.empty:
            df_vacas.to_excel(writer, sheet_name='Vacas', index=False)
        if not df_estacoes_global.empty:
            df_estacoes_global.to_excel(writer, sheet_name='Estacoes_Global', index=False)
        if not df_estacoes_categoria.empty:
            df_estacoes_categoria.to_excel(writer, sheet_name='Estacoes_por_Categoria', index=False)
        if not df_estacoes_protocolo.empty:
            df_estacoes_protocolo.to_excel(writer, sheet_name='Estacoes_por_Protocolo', index=False)

        workbook = writer.book
        percent_format = workbook.add_format({'num_format': '0.0%'}) 
        decimal_format = workbook.add_format({'num_format': '0.00'}) 
        date_format = workbook.add_format({'num_format': 'dd-mm-yyyy'})
        integer_format = workbook.add_format({'num_format': '0'}) # Formato para inteiros

        abas_dfs = {
            'Vacas': df_vacas,
            'Estacoes_Global': df_estacoes_global,
            'Estacoes_por_Categoria': df_estacoes_categoria,
            'Estacoes_por_Protocolo': df_estacoes_protocolo
        }

        # Aplica formatação coluna por coluna
        for nome_aba, df_aba in abas_dfs.items():
            if df_aba.empty: continue
            worksheet = writer.sheets[nome_aba]
            for i, col_name in enumerate(df_aba.columns):
                col_name_lower = col_name.lower()
                # Porcentagem
                if 'taxa' in col_name_lower or 'concepcao_por_estação' in col_name_lower or 'protocolos' in col_name_lower and 'total' not in col_name_lower:
                     worksheet.set_column(i, i, 15, percent_format)
                # Decimal (Peso)
                elif 'peso_medio' in col_name_lower:
                     worksheet.set_column(i, i, 18, decimal_format)
                # Data
                elif 'data_media' in col_name_lower:
                     worksheet.set_column(i, i, 15, date_format)
                # Inteiro (Contagens)
                elif 'total' in col_name_lower or 'qtd_' in col_name_lower or 'numero_de' in col_name_lower or '_abortos' in col_name_lower or '_prenhezes' in col_name_lower:
                     worksheet.set_column(i, i, 12, integer_format)
                 # Ajuste de Largura Padrão
                elif 'animal_id' in col_name_lower:
                     worksheet.set_column(i, i, 15)
                elif 'categoria' in col_name_lower:
                     worksheet.set_column(i, i, 20)
                else: # Default para outras colunas
                    worksheet.set_column(i, i, 12)


    print("\nArquivo 'saida_analise.xlsx' gerado com sucesso!")
    print(f"A análise considerou as abas: {abas_selecionadas}")
except Exception as e:
    print(f"[ERRO] Não foi possível salvar o arquivo 'saida_analise.xlsx'. Verifique se ele está aberto ou se há permissão de escrita. Erro: {e}")


# --- Seção de Geração de Gráficos ---
# (Permanece igual, mas atenção: novas colunas de taxa precisam ser numéricas no DF para plotar)
if not df_estacoes_global.empty:
    output_dir_base = "graficos"
    output_dir_geral = os.path.join(output_dir_base, "geral")
    output_dir_categorias = os.path.join(output_dir_base, "categorias")

    try:
        if not os.path.exists(output_dir_geral): os.makedirs(output_dir_geral)
        if not os.path.exists(output_dir_categorias): os.makedirs(output_dir_categorias)
        print(f"Gráficos serão salvos nas subpastas dentro de: '{output_dir_base}'")

        def criar_grafico_barras(x_data, y_data, titulo, eixo_y_label, caminho_arquivo, is_percent=False):
            plt.figure(figsize=(12, 7))
            y_data_numeric = pd.to_numeric(y_data, errors='coerce')
            
            # Ajuste: Se for %, os dados já estão em formato decimal (0.55), não precisa dividir
            
            bars = plt.bar(x_data.astype(str), y_data_numeric, color='skyblue')
            plt.title(titulo, fontsize=16)
            plt.ylabel(eixo_y_label, fontsize=12)
            plt.xticks(rotation=45, ha='right')
            
            if is_percent:
                from matplotlib.ticker import PercentFormatter
                plt.gca().yaxis.set_major_formatter(PercentFormatter(xmax=1.0)) 
                
            for i, bar in enumerate(bars):
                label_val = y_data_numeric.iloc[i]
                if pd.isna(label_val): continue
                
                if is_percent:
                    label = f'{label_val*100:.1f}%' 
                else:
                    try: label = f'{int(label_val)}'
                    except: label = f'{label_val:.2f}'

                plt.text(bar.get_x() + bar.get_width()/2.0, bar.get_height(), label, va='bottom', ha='center')
                
            plt.tight_layout()
            plt.savefig(caminho_arquivo)
            plt.close()

        df_global_sorted = df_estacoes_global.sort_values('estacao')
        # Adiciona a nova taxa geral de protocolos aos gráficos
        graficos_globais = [
            ('total_concepcoes', 'Total de Concepções por Estação', 'Nº de Concepções'),
            ('total_registros', 'Total de Registros por Estação', 'Nº de Registros'),
            ('total_abortos', 'Total de Abortos por Estação', 'Nº de Abortos'),
            ('taxa_prenhez_geral', 'Taxa de Prenhez Geral por Estação', 'Taxa de Prenhez'),
            ('total_protocolos_aplicados', 'Total Protocolos Aplicados por Estação', 'Nº Protocolos'), # Novo
            ('total_protocolos_sucesso', 'Total Protocolos Sucesso por Estação', 'Nº Sucessos'),     # Novo
            ('taxa_geral_protocolos', 'Taxa Geral Sucesso Protocolos por Estação', 'Taxa de Sucesso (%)') 
        ]
        
        print("Gerando gráficos...")
        for coluna, titulo, eixo_y in tqdm(graficos_globais, desc="Gerando Gráficos Gerais"):
            # Verifica se a coluna existe no DataFrame antes de tentar plotar
            if coluna not in df_global_sorted.columns:
                 print(f"[AVISO] Coluna '{coluna}' não encontrada para gráfico geral. Pulando.")
                 continue

            caminho = os.path.join(output_dir_geral, f"geral_{coluna}.png")
            # Ajusta a condição para is_percent
            is_percent = 'taxa' in coluna or 'concepcao' in coluna or ('protocolos' in coluna and 'total' not in coluna)
            df_plot = df_global_sorted.dropna(subset=[coluna])
            if not df_plot.empty:
                 try:
                     y_values = pd.to_numeric(df_plot[coluna], errors='coerce')
                     if y_values.notna().any(): 
                         criar_grafico_barras(df_plot['estacao'], y_values, titulo, eixo_y, caminho, is_percent=is_percent)
                     #else: print(f"[AVISO] Sem dados numéricos válidos para plotar gráfico: {titulo}") # Debug
                 except Exception as e:
                      print(f"[ERRO] Falha ao gerar gráfico '{titulo}': {e}")
        
        # Adiciona a nova taxa geral de protocolos aos gráficos por categoria
        anos = df_estacoes_categoria['estacao'].unique()
        graficos_categoria = [
            ('total_concepcoes_cat', 'Total de Concepções por Categoria', 'Nº de Concepções'),
            ('total_registros_cat', 'Total de Registros por Categoria', 'Nº de Registros'),
            ('total_abortos_cat', 'Total de Abortos por Categoria', 'Nº de Abortos'),
            ('taxa_prenhez_cat', 'Taxa de Prenhez por Categoria', 'Taxa de Prenhez'),
            ('total_protocolos_aplicados_cat', 'Total Protocolos Aplicados por Categoria', 'Nº Protocolos'), # Novo
            ('total_protocolos_sucesso_cat', 'Total Protocolos Sucesso por Categoria', 'Nº Sucessos'),     # Novo
            ('taxa_geral_protocolos_cat', 'Taxa Geral Sucesso Protocolos por Categoria', 'Taxa de Sucesso (%)')
        ]
        for ano in tqdm(anos, desc="Gerando Gráficos por Categoria"):
            df_ano = df_estacoes_categoria[df_estacoes_categoria['estacao'] == ano].sort_values('categoria')
            for coluna, titulo_base, eixo_y in graficos_categoria:
                 if coluna not in df_ano.columns:
                      #print(f"[AVISO] Coluna '{coluna}' não encontrada para gráfico categoria {ano}. Pulando.") # Debug
                      continue

                 titulo_completo = f'{titulo_base} - {ano}'
                 caminho = os.path.join(output_dir_categorias, f"categoria_{ano}_{coluna}.png")
                 is_percent = 'taxa' in coluna or 'concepcao' in coluna or ('protocolos' in coluna and 'total' not in coluna)
                 df_plot = df_ano.dropna(subset=[coluna])
                 if not df_plot.empty:
                    try:
                        y_values = pd.to_numeric(df_plot[coluna], errors='coerce')
                        if y_values.notna().any():
                             criar_grafico_barras(df_plot['categoria'], y_values, titulo_completo, eixo_y, caminho, is_percent=is_percent)
                        #else: print(f"[AVISO] Sem dados numéricos válidos para plotar gráfico: {titulo_completo}") # Debug
                    except Exception as e:
                        print(f"[ERRO] Falha ao gerar gráfico '{titulo_completo}': {e}")
                        
    except Exception as e:
        print(f"[ERRO] Ocorreu um erro durante a geração dos gráficos: {e}")


# --- Exibição de informações no console ---
print("\nHistóricos inválidos encontrados (ignorados na análise):")
if blacklist_historicos:
    for h in sorted(list(blacklist_historicos)):
        print(f"  - {h}")
else:
    print("  Nenhum histórico inválido encontrado.")
print(f"\nTotal de vacas ignoradas na análise individual (por terem histórico inválido em algum momento): {num_animais_ignorados}")