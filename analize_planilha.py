
import pandas as pd
import matplotlib.pyplot as plt


# Carrega o arquivo Excel
xl_file = pd.ExcelFile('CONTROLE ESTACAO.xlsx')


# Lê todas as abas da planilha por índice, ignorando abas problemáticas
dfs = {}
abas_com_erro = []
for i in range(len(xl_file.sheet_names)):
    nome_aba = xl_file.sheet_names[i]
    try:
        df = xl_file.parse(i)
        dfs[i] = df  # usa o índice como chave
    except Exception as e:
        print(f"[AVISO] Erro ao ler a aba de índice {i} ('{nome_aba}'): {e}")
        abas_com_erro.append((i, nome_aba))
if abas_com_erro:
    indices = [str(idx) for idx, _ in abas_com_erro]
    nomes = [nome for _, nome in abas_com_erro]
    print(f"\nAs seguintes abas não puderam ser lidas e foram ignoradas (para corigir desligue os filtros):")
    for idx, nome in abas_com_erro:
        print(f"  Índice: {idx} | Nome: {nome}")


# === Busca de múltiplos animais específicos em todas as abas ===
# Agora os IDs são extraídos automaticamente da última estação (última aba da planilha)
ultima_aba = xl_file.sheet_names[-1]
df_full = xl_file.parse(ultima_aba)
header_row = df_full.iloc[0]
animal_idx = None
for idx, col_name in enumerate(header_row):
    nome = str(col_name).strip().upper()
    if nome == 'ANIMAL':
        animal_idx = idx
        break
if animal_idx is None:
    raise ValueError("Coluna 'ANIMAL' não encontrada na última estação.")
df = df_full.iloc[1:].reset_index(drop=True)
animal_ids = df.iloc[:, animal_idx].astype(str).str.strip().unique().tolist()
animal_ids = [aid for aid in animal_ids if aid]


vacas_data = []
for animal_id in animal_ids:
    pesos_por_ano = {}
    prenhez_por_ano = {}

    for aba in xl_file.sheet_names:
        df_full = xl_file.parse(aba)
        header_row = df_full.iloc[0]
        historico_idx = None
        situacao_idx = None
        categoria_idx = None
        peso_idx = None
        animal_idx = None
        for idx, col_name in enumerate(header_row):
            nome = str(col_name).strip().upper()
            if nome == 'HISTÓRICO':
                historico_idx = idx
            if nome == 'SITUAÇÃO':
                situacao_idx = idx
            if nome == 'CATEGORIA':
                categoria_idx = idx
            if nome in 'PESO 205':
                peso_idx = idx
            if nome == 'ANIMAL':
                animal_idx = idx
        if situacao_idx is None or peso_idx is None or animal_idx is None:
            continue
        df = df_full.iloc[1:].reset_index(drop=True)
        # Busca animal na coluna 'ANIMAL'
        mask = df.iloc[:, animal_idx].astype(str).str.strip() == animal_id
        df_animal = df[mask]
        if not df_animal.empty:
            # Peso de desmame
            pesos = pd.to_numeric(df_animal.iloc[:, peso_idx], errors='coerce').dropna()
            if not pesos.empty:
                pesos_por_ano[aba] = list(pesos)
            # Taxa de prenhez
            situacao = df_animal.iloc[:, situacao_idx].astype(str)
            prenhez = (situacao.str.contains('P') | situacao.str.contains('AB')).sum()
            total = len(situacao)
            taxa = prenhez / total if total > 0 else 0
            prenhez_por_ano[aba] = taxa

    # Consolidar dados de todas as abas para cálculo global
    total_tentativas = 0
    total_prenhezes = 0
    total_abortos = 0
    todos_pesos = []
    pesos_machos = []
    pesos_femeas = []
    qtd_machos = 0
    qtd_femeas = 0
    for ano, taxa in prenhez_por_ano.items():
        df_full = xl_file.parse(ano)
        header_row = df_full.iloc[0]
        animal_idx = None
        situacao_idx = None
        peso_idx = None
        sexo_idx = None
        for idx, col_name in enumerate(header_row):
            nome = str(col_name).strip().upper()
            if nome == 'ANIMAL':
                animal_idx = idx
            if nome == 'SITUAÇÃO':
                situacao_idx = idx
            if nome == "PESO 205":
                peso_idx = idx
            if nome == "SEXO":
                sexo_idx = idx
        if animal_idx is None or situacao_idx is None:
            continue
        df = df_full.iloc[1:].reset_index(drop=True)
        mask = df.iloc[:, animal_idx].astype(str).str.strip() == animal_id
        df_animal = df[mask]
        situacoes = df_animal.iloc[:, situacao_idx].astype(str)
        tentativas = len(situacoes)
        prenhezes = (situacoes.str.contains('P') | situacoes.str.contains('AB')).sum()
        abortos = situacoes.str.contains('AB').sum()
        total_tentativas += tentativas
        total_prenhezes += prenhezes
        total_abortos += abortos
        if sexo_idx is not None:
            sexos = df_animal.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
            qtd_machos += (sexos == 'M').sum()
            qtd_femeas += (sexos == 'F').sum()
        if peso_idx is not None:
            pesos = pd.to_numeric(df_animal.iloc[:, peso_idx], errors='coerce').dropna()
            todos_pesos.extend(list(pesos))
            if sexo_idx is not None:
                sexos_pesos = df_animal.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
                sexos_validos = sexos_pesos[pesos.index]
                pesos_machos.extend(pesos[sexos_validos == 'M'].tolist())
                pesos_femeas.extend(pesos[sexos_validos == 'F'].tolist())

    prot_labels = [f"{i}PROT" for i in range(1, 15)]
    prot_total = 0
    prot_prenhez = 0
    prot_preenhez_individual = {prot: 0 for prot in prot_labels}
    prot_total_individual = {prot: 0 for prot in prot_labels}
    for ano in prenhez_por_ano.keys():
        df_full = xl_file.parse(ano)
        header_row = df_full.iloc[0]
        animal_idx = None
        historico_idx = None
        for idx, col_name in enumerate(header_row):
            nome = str(col_name).strip().upper()
            if nome == 'ANIMAL':
                animal_idx = idx
            if nome == 'HISTÓRICO':
                historico_idx = idx
        if animal_idx is None or historico_idx is None:
            continue
        df = df_full.iloc[1:].reset_index(drop=True)
        mask = df.iloc[:, animal_idx].astype(str).str.strip() == animal_id
        df_animal = df[mask]
        for _, row in df_animal.iterrows():
            historico = str(row.iloc[historico_idx])
            for prot in prot_labels:
                if f"{prot}-P" in historico:
                    prot_prenhez += 1
                    prot_total += 1
                    prot_preenhez_individual[prot] += 1
                    prot_total_individual[prot] += 1
                elif f"{prot}-R" in historico:
                    prot_total += 1
                    prot_total_individual[prot] += 1

    porcentagem_prot_prenhez = 100 * prot_prenhez / prot_total if prot_total > 0 else None
    taxas_prot_individual = {}
    for prot in prot_labels:
        total = prot_total_individual[prot]
        prenhez = prot_preenhez_individual[prot]
        taxa = prenhez / total if total > 0 else None
        taxas_prot_individual[prot] = (taxa, prenhez, total)

    vacas_data.append({
        'animal_id': animal_id,
        'total_tentativas': total_tentativas,
        'total_prenhezes': total_prenhezes,
        'total_abortos': total_abortos,
        'num_estacoes': len(prenhez_por_ano),
        'taxa_global': total_prenhezes / total_tentativas if total_tentativas > 0 else None,
        'peso_medio_global': sum(todos_pesos)/len(todos_pesos) if todos_pesos else None,
        'peso_medio_machos': sum(pesos_machos)/len(pesos_machos) if pesos_machos else None,
        'peso_medio_femeas': sum(pesos_femeas)/len(pesos_femeas) if pesos_femeas else None,
        'qtd_machos': qtd_machos,
        'qtd_femeas': qtd_femeas,
        'porcentagem_prot_prenhez': porcentagem_prot_prenhez,
        'taxas_prot_individual': taxas_prot_individual
    })

    # === Resultados detalhados para o animal buscado ===
    print(f"\nResultados para o animal '{animal_id}':")

    # Consolidar dados de todas as abas para cálculo global
    total_tentativas = 0
    total_prenhezes = 0
    total_abortos = 0
    todos_pesos = []
    pesos_machos = []
    pesos_femeas = []
    qtd_machos = 0
    qtd_femeas = 0
    for ano, taxa in prenhez_por_ano.items():
        # Para cada aba, recontar tentativas, prenhezes e abortos
        df_full = xl_file.parse(ano)
        header_row = df_full.iloc[0]
        animal_idx = None
        situacao_idx = None
        peso_idx = None
        sexo_idx = None
        for idx, col_name in enumerate(header_row):
            nome = str(col_name).strip().upper()
            if nome == 'ANIMAL':
                animal_idx = idx
            if nome == 'SITUAÇÃO':
                situacao_idx = idx
            if nome == "PESO 205":
                peso_idx = idx
            if nome == "SEXO":
                sexo_idx = idx
        if animal_idx is None or situacao_idx is None:
            continue
        df = df_full.iloc[1:].reset_index(drop=True)
        mask = df.iloc[:, animal_idx].astype(str).str.strip() == animal_id
        df_animal = df[mask]
        situacoes = df_animal.iloc[:, situacao_idx].astype(str)
        tentativas = len(situacoes)
        prenhezes = (situacoes.str.contains('P') | situacoes.str.contains('AB')).sum()
        abortos = situacoes.str.contains('AB').sum()
        total_tentativas += tentativas
        total_prenhezes += prenhezes
        total_abortos += abortos
        if sexo_idx is not None:
            sexos = df_animal.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
            qtd_machos += (sexos == 'M').sum()
            qtd_femeas += (sexos == 'F').sum()
        if peso_idx is not None:
            pesos = pd.to_numeric(df_animal.iloc[:, peso_idx], errors='coerce').dropna()
            todos_pesos.extend(list(pesos))
            # Peso médio de desmame por sexo (apenas onde há peso)
            if sexo_idx is not None:
                sexos_pesos = df_animal.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
                # Alinha sexos e pesos apenas para linhas com peso válido
                sexos_validos = sexos_pesos[pesos.index]
                pesos_machos.extend(pesos[sexos_validos == 'M'].tolist())
                pesos_femeas.extend(pesos[sexos_validos == 'F'].tolist())

    # Exibe por estação
    if pesos_por_ano:
        for ano, pesos in pesos_por_ano.items():
            print(f"Ano/Estação: {ano} | Pesos de desmame: {pesos}")
        anos = list(pesos_por_ano.keys())
        medias = [sum(pesos)/len(pesos) for pesos in pesos_por_ano.values()]
        # TA COMENTADO
        #plt.figure()
        #plt.bar(anos, medias)
        #plt.ylabel('Peso médio de desmame')
        #plt.xlabel('Ano/Estação')
        #plt.title(f'Peso médio de desmame dos bezerros do animal {animal_id}')
        #plt.show()
    else:
        print('Nenhum peso de desmame encontrado para este animal.')
    if prenhez_por_ano:
        for ano, taxa in prenhez_por_ano.items():
            print(f"Ano/Estação: {ano} | Taxa de prenhez: {taxa:.2%}")
    else:
        print('Nenhuma informação de prenhez encontrada para este animal.')

    # Exibe consolidado global
    num_estacoes = len(prenhez_por_ano)
    if total_tentativas > 0:
        taxa_global = total_prenhezes / total_tentativas
        print(f"\n[Consolidado Global] Total de tentativas: {total_tentativas} | Total de prenhezes: {total_prenhezes} | Total de abortos: {total_abortos} | Número de estações: {num_estacoes} | Taxa global de concepção: {taxa_global:.2%}")
    else:
        print(f"\n[Consolidado Global] Nenhuma tentativa encontrada para este animal. Número de estações: {num_estacoes}")
    if todos_pesos:
        print(f"Peso médio de desmame global: {sum(todos_pesos)/len(todos_pesos):.2f}")
        if pesos_machos:
            print(f"Peso médio de desmame dos bezerros machos: {sum(pesos_machos)/len(pesos_machos):.2f}")
        else:
            print("Peso médio de desmame dos bezerros machos: Não encontrado.")
        if pesos_femeas:
            print(f"Peso médio de desmame das bezerras fêmeas: {sum(pesos_femeas)/len(pesos_femeas):.2f}")
        else:
            print("Peso médio de desmame das bezerras fêmeas: Não encontrado.")
    else:
        print("Peso médio de desmame global: Não encontrado.")

    print(f"Quantidade de bezerros machos: {qtd_machos}")
    print(f"Quantidade de bezerras fêmeas: {qtd_femeas}")

    # --- Porcentagem de PROTs com prenhez para a vaca e taxa de concepção por protocolo (PROT) individualmente ---
    prot_labels = [f"{i}PROT" for i in range(1, 15)]
    prot_total = 0
    prot_prenhez = 0
    prot_preenhez_individual = {prot: 0 for prot in prot_labels}
    prot_total_individual = {prot: 0 for prot in prot_labels}
    for ano in prenhez_por_ano.keys():
        df_full = xl_file.parse(ano)
        header_row = df_full.iloc[0]
        animal_idx = None
        historico_idx = None
        for idx, col_name in enumerate(header_row):
            nome = str(col_name).strip().upper()
            if nome == 'ANIMAL':
                animal_idx = idx
            if nome == 'HISTÓRICO':
                historico_idx = idx
        if animal_idx is None or historico_idx is None:
            continue
        df = df_full.iloc[1:].reset_index(drop=True)
        mask = df.iloc[:, animal_idx].astype(str).str.strip() == animal_id
        df_animal = df[mask]
        for _, row in df_animal.iterrows():
            historico = str(row.iloc[historico_idx])
            for prot in prot_labels:
                if f"{prot}-P" in historico:
                    prot_prenhez += 1
                    prot_total += 1
                    prot_preenhez_individual[prot] += 1
                    prot_total_individual[prot] += 1
                elif f"{prot}-R" in historico:
                    prot_total += 1
                    prot_total_individual[prot] += 1
    if prot_total > 0:
        porcentagem_prot_prenhez = 100 * prot_prenhez / prot_total
        print(f"Porcentagem de protocolos (PROT) com prenhez para a vaca: {porcentagem_prot_prenhez:.2f}% ({prot_prenhez}/{prot_total})")
        print("\nTaxa de concepção por protocolo (PROT) para a vaca específica:")
        for prot in prot_labels:
            total = prot_total_individual[prot]
            prenhez = prot_preenhez_individual[prot]
            taxa = prenhez / total if total > 0 else 0
            print(f"Protocolo: {prot} | Taxa de prenhez: {taxa:.2%} ({prenhez}/{total})")
    else:
        print("A vaca não participou de nenhum protocolo PROT.")


# === Resultados detalhados para o animal buscado ===
print(f"\nResultados para o animal '{animal_id}':")

 # Consolidar dados de todas as abas para cálculo global
total_tentativas = 0
total_prenhezes = 0
total_abortos = 0
todos_pesos = []
pesos_machos = []
pesos_femeas = []
qtd_machos = 0
qtd_femeas = 0
for ano, taxa in prenhez_por_ano.items():
    # Para cada aba, recontar tentativas, prenhezes e abortos
    df_full = xl_file.parse(ano)
    header_row = df_full.iloc[0]
    animal_idx = None
    situacao_idx = None
    peso_idx = None
    sexo_idx = None
    for idx, col_name in enumerate(header_row):
        nome = str(col_name).strip().upper()
        if nome == 'ANIMAL':
            animal_idx = idx
        if nome == 'SITUAÇÃO':
            situacao_idx = idx
        if nome == "PESO 205":
            peso_idx = idx
        if nome == "SEXO":
            sexo_idx = idx
    if animal_idx is None or situacao_idx is None:
        continue
    df = df_full.iloc[1:].reset_index(drop=True)
    mask = df.iloc[:, animal_idx].astype(str).str.strip() == animal_id
    df_animal = df[mask]
    situacoes = df_animal.iloc[:, situacao_idx].astype(str)
    tentativas = len(situacoes)
    prenhezes = (situacoes.str.contains('P') | situacoes.str.contains('AB')).sum()
    abortos = situacoes.str.contains('AB').sum()
    total_tentativas += tentativas
    total_prenhezes += prenhezes
    total_abortos += abortos
    if sexo_idx is not None:
        sexos = df_animal.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
        qtd_machos += (sexos == 'M').sum()
        qtd_femeas += (sexos == 'F').sum()
    if peso_idx is not None:
        pesos = pd.to_numeric(df_animal.iloc[:, peso_idx], errors='coerce').dropna()
        todos_pesos.extend(list(pesos))
        # Peso médio de desmame por sexo (apenas onde há peso)
        if sexo_idx is not None:
            sexos_pesos = df_animal.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
            # Alinha sexos e pesos apenas para linhas com peso válido
            sexos_validos = sexos_pesos[pesos.index]
            pesos_machos.extend(pesos[sexos_validos == 'M'].tolist())
            pesos_femeas.extend(pesos[sexos_validos == 'F'].tolist())

# Exibe por estação
if pesos_por_ano:
    for ano, pesos in pesos_por_ano.items():
        print(f"Ano/Estação: {ano} | Pesos de desmame: {pesos}")
    anos = list(pesos_por_ano.keys())
    medias = [sum(pesos)/len(pesos) for pesos in pesos_por_ano.values()]
    # TA COMENTADO
    #plt.figure()
    #plt.bar(anos, medias)
    #plt.ylabel('Peso médio de desmame')
    #plt.xlabel('Ano/Estação')
    #plt.title(f'Peso médio de desmame dos bezerros do animal {animal_id}')
    #plt.show()
else:
    print('Nenhum peso de desmame encontrado para este animal.')
if prenhez_por_ano:
    for ano, taxa in prenhez_por_ano.items():
        print(f"Ano/Estação: {ano} | Taxa de prenhez: {taxa:.2%}")
else:
    print('Nenhuma informação de prenhez encontrada para este animal.')

# Exibe consolidado global
num_estacoes = len(prenhez_por_ano)
if total_tentativas > 0:
    taxa_global = total_prenhezes / total_tentativas
    print(f"\n[Consolidado Global] Total de tentativas: {total_tentativas} | Total de prenhezes: {total_prenhezes} | Total de abortos: {total_abortos} | Número de estações: {num_estacoes} | Taxa global de concepção: {taxa_global:.2%}")
else:
    print(f"\n[Consolidado Global] Nenhuma tentativa encontrada para este animal. Número de estações: {num_estacoes}")
if todos_pesos:
    print(f"Peso médio de desmame global: {sum(todos_pesos)/len(todos_pesos):.2f}")
    if pesos_machos:
        print(f"Peso médio de desmame dos bezerros machos: {sum(pesos_machos)/len(pesos_machos):.2f}")
    else:
        print("Peso médio de desmame dos bezerros machos: Não encontrado.")
    if pesos_femeas:
        print(f"Peso médio de desmame das bezerras fêmeas: {sum(pesos_femeas)/len(pesos_femeas):.2f}")
    else:
        print("Peso médio de desmame das bezerras fêmeas: Não encontrado.")
else:
    print("Peso médio de desmame global: Não encontrado.")

print(f"Quantidade de bezerros machos: {qtd_machos}")
print(f"Quantidade de bezerras fêmeas: {qtd_femeas}")

# --- Porcentagem de PROTs com prenhez para a vaca e taxa de concepção por protocolo (PROT) individualmente ---
prot_labels = [f"{i}PROT" for i in range(1, 15)]
prot_total = 0
prot_prenhez = 0
prot_preenhez_individual = {prot: 0 for prot in prot_labels}
prot_total_individual = {prot: 0 for prot in prot_labels}
for ano in prenhez_por_ano.keys():
    df_full = xl_file.parse(ano)
    header_row = df_full.iloc[0]
    animal_idx = None
    historico_idx = None
    for idx, col_name in enumerate(header_row):
        nome = str(col_name).strip().upper()
        if nome == 'ANIMAL':
            animal_idx = idx
        if nome == 'HISTÓRICO':
            historico_idx = idx
    if animal_idx is None or historico_idx is None:
        continue
    df = df_full.iloc[1:].reset_index(drop=True)
    mask = df.iloc[:, animal_idx].astype(str).str.strip() == animal_id
    df_animal = df[mask]
    for _, row in df_animal.iterrows():
        historico = str(row.iloc[historico_idx])
        for prot in prot_labels:
            if f"{prot}-P" in historico:
                prot_prenhez += 1
                prot_total += 1
                prot_preenhez_individual[prot] += 1
                prot_total_individual[prot] += 1
            elif f"{prot}-R" in historico:
                prot_total += 1
                prot_total_individual[prot] += 1
if prot_total > 0:
    porcentagem_prot_prenhez = 100 * prot_prenhez / prot_total
    print(f"Porcentagem de protocolos (PROT) com prenhez para a vaca: {porcentagem_prot_prenhez:.2f}% ({prot_prenhez}/{prot_total})")
    print("\nTaxa de concepção por protocolo (PROT) para a vaca específica:")
    for prot in prot_labels:
        total = prot_total_individual[prot]
        prenhez = prot_preenhez_individual[prot]
        taxa = prenhez / total if total > 0 else 0
        print(f"Protocolo: {prot} | Taxa de prenhez: {taxa:.2%} ({prenhez}/{total})")
else:
    print("A vaca não participou de nenhum protocolo PROT.")

# === Fim da busca por animal específico ===

# Usa a primeira aba da planilha e ignora a primeira linha (header manual)

# === Bloco unificado e limpo para métricas principais ===
df_full = list(dfs.values())[0]
header_row = df_full.iloc[0]

# Busca dinâmica dos índices das colunas principais
def get_col_idx(header_row, col_name):
    for idx, col in enumerate(header_row):
        nome = str(col).strip().upper()
        for target in col_name:
            if nome == target:
                return idx
    return None

historico_idx = get_col_idx(header_row, ['HISTÓRICO'])
situacao_idx = get_col_idx(header_row, ['SITUAÇÃO'])
categoria_idx = get_col_idx(header_row, ['CATEGORIA'])
peso_idx = get_col_idx(header_row, ['PESO 205'])


if historico_idx is None or situacao_idx is None or categoria_idx is None or peso_idx is None:
    raise ValueError("Colunas 'HISTÓRICO', 'SITUAÇÃO', 'PESO 205' ou 'CATEGORIA' não encontradas na segunda linha da planilha.")



# Sempre defina df_filtrado corretamente para o DataFrame de análise
df_filtrado = df  # Se quiser filtrar por peso, pode ajustar aqui
situacao = df_filtrado.iloc[:, situacao_idx].astype(str)

# Taxa de prenhez por categoria
print(f"\nTaxa de prenhez por categoria de vaca:")
categorias = df_filtrado.iloc[:, categoria_idx].unique()
taxas = []
prenhezes = []
totais = []
nomes_categorias = []
for cat in categorias:
    if pd.isna(cat):
        continue
    df_cat = df_filtrado[df_filtrado.iloc[:, categoria_idx] == cat]
    situacao_cat = df_cat.iloc[:, situacao_idx].astype(str)
    total_cat = len(df_cat)
    prenhezes_cat = situacao_cat.str.contains('P').sum() + situacao_cat.str.contains('AB').sum()
    taxa_cat = prenhezes_cat / total_cat if total_cat > 0 else 0
    taxas.append(taxa_cat)
    prenhezes.append(prenhezes_cat)
    totais.append(total_cat)
    nomes_categorias.append(str(cat))
    print(f"Categoria: {cat} | Taxa de prenhez: {taxa_cat:.2%} ({prenhezes_cat}/{total_cat})")

# Peso médio de desmame por categoria
print("\nPeso médio de desmame por categoria:")
pesos_medios = []
nomes_categorias_peso = []
if peso_idx is not None:
    for cat in categorias:
        if pd.isna(cat):
            continue
        df_cat = df_filtrado[df_filtrado.iloc[:, categoria_idx] == cat]
        pesos_cat = pd.to_numeric(df_cat.iloc[:, peso_idx], errors='coerce')
        media_peso_cat = pesos_cat.mean()
        pesos_medios.append(media_peso_cat)
        nomes_categorias_peso.append(str(cat))
        print(f"Categoria: {cat} | Peso médio de desmame: {media_peso_cat:.2f}")
    pesos_global = pd.to_numeric(df_filtrado.iloc[:, peso_idx], errors='coerce')
    media_peso_global = pesos_global.mean()
    print(f"\nPeso médio de desmame global: {media_peso_global:.2f}")
else:
    print("[AVISO] Não foi possível calcular peso médio de desmame por categoria/global nesta aba.")

# Taxa global de concepção por protocolo
print("\nTaxa global de concepção por protocolo:")
protocolos = df_filtrado.iloc[:, historico_idx]
total_protocolos = protocolos.notna().sum()
# Para concepções, considerar coluna SITUAÇÃO com 'P' ou 'AB'
concepcoes_protocolos = situacao.str.contains('P').sum() + situacao.str.contains('AB').sum()
taxa_global_protocolo = concepcoes_protocolos / total_protocolos if total_protocolos > 0 else 0
print(f"Taxa global de concepção: {taxa_global_protocolo:.2%} ({concepcoes_protocolos}/{total_protocolos})")

# Taxa de concepção por protocolo (apenas protocolos com '-P' ou '-R')
print("\nTaxa de concepção por protocolo (apenas protocolos com '-P' ou '-R'):")
prot_labels = [f"{i}PROT" for i in range(1, 15)]
prot_preenhez = {prot: 0 for prot in prot_labels}
prot_total = {prot: 0 for prot in prot_labels}
for idx, row in df_filtrado.iterrows():
    historico = str(row.iloc[historico_idx])
    if pd.isna(historico) or not historico.strip():
        continue
    for prot in prot_labels:
        if f"{prot}-P" in historico:
            prot_preenhez[prot] += 1
            prot_total[prot] += 1
        elif f"{prot}-R" in historico:
            prot_total[prot] += 1
taxas_prot = []
nomes_prot = []
for prot in prot_labels:
    total = prot_total[prot]
    prenhezes = prot_preenhez[prot]
    taxa = prenhezes / total if total > 0 else 0
    print(f"Protocolo: {prot} | Taxa de prenhez: {taxa:.2%} ({prenhezes}/{total})")
    nomes_prot.append(prot)
    taxas_prot.append(taxa)

# Contagem de vacas prenhas e abortos
total_prenhas = situacao.str.contains('P').sum() + situacao.str.contains('AB').sum()
total_abortos = situacao.str.contains('AB').sum()
print(f"\nQuantidade de vacas prenhas (incluindo abortos): {total_prenhas}")
print(f"Quantidade de abortos na estação: {total_abortos}")



estacoes_data = []
for aba in xl_file.sheet_names:
    try:
        df_full = xl_file.parse(aba)
        header_row = df_full.iloc[0]
        def get_col_idx(header_row, col_name):
            for idx, col in enumerate(header_row):
                nome = str(col).strip().upper()
                for target in col_name:
                    if nome == target:
                        return idx
            return None

        historico_idx = get_col_idx(header_row, ['HISTÓRICO'])
        situacao_idx = get_col_idx(header_row, ['SITUAÇÃO'])
        categoria_idx = get_col_idx(header_row, ['CATEGORIA'])
        if historico_idx is None or situacao_idx is None or categoria_idx is None:
            continue
        df = df_full.iloc[1:].reset_index(drop=True)
        situacao = df.iloc[:, situacao_idx].astype(str)
        categorias = df.iloc[:, categoria_idx].unique()

        total_tentativas = len(df)
        total_concepcoes = situacao.str.contains('P').sum() + situacao.str.contains('AB').sum()
        taxa_global = total_concepcoes / total_tentativas if total_tentativas > 0 else None

        taxas_categoria = {}
        for cat in categorias:
            if pd.isna(cat):
                continue
            df_cat = df[df.iloc[:, categoria_idx] == cat]
            situacao_cat = df_cat.iloc[:, situacao_idx].astype(str)
            total_cat = len(df_cat)
            concepcoes_cat = situacao_cat.str.contains('P').sum() + situacao_cat.str.contains('AB').sum()
            taxa_cat = concepcoes_cat / total_cat if total_cat > 0 else None
            taxas_categoria[str(cat)] = (taxa_cat, concepcoes_cat, total_cat)

        prot_labels = [f"{i}PROT" for i in range(1, 15)]
        prot_preenhez = {prot: 0 for prot in prot_labels}
        prot_total = {prot: 0 for prot in prot_labels}
        for idx, row in df.iterrows():
            historico = str(row.iloc[historico_idx])
            if pd.isna(historico) or not historico.strip():
                continue
            for prot in prot_labels:
                if f"{prot}-P" in historico:
                    prot_preenhez[prot] += 1
                    prot_total[prot] += 1
                elif f"{prot}-R" in historico:
                    prot_total[prot] += 1

        taxas_prot = {}
        for prot in prot_labels:
            total = prot_total[prot]
            prenhez = prot_preenhez[prot]
            taxa = prenhez / total if total > 0 else None
            taxas_prot[prot] = (taxa, prenhez, total)

        data_ia_idx = get_col_idx(header_row, ['DATA IA'])
        data_ia_med = None
        if data_ia_idx is not None:
            datas_ia = pd.to_datetime(df.iloc[:, data_ia_idx], errors='coerce')
            mask_prenhez = situacao.str.contains('P') | situacao.str.contains('AB')
            datas_ia_prenhez = datas_ia[mask_prenhez]
            if not datas_ia_prenhez.dropna().empty:
                data_ia_med = datas_ia_prenhez.dropna().mean()

        estacoes_data.append({
            'estacao': aba,
            'total_tentativas': total_tentativas,
            'total_concepcoes': total_concepcoes,
            'taxa_global': taxa_global,
            'taxas_categoria': taxas_categoria,
            'taxas_prot': taxas_prot,
            'data_ia_media': data_ia_med.strftime('%d/%m/%Y') if data_ia_med is not None else None
        })
    except Exception as e:
        continue

# === Exporta para Excel ===
import numpy as np
import pandas as pd
import xlsxwriter

# Página 1: vacas
vacas_rows = []
for vaca in vacas_data:
    row = {
        'animal_id': vaca['animal_id'],
        'total_tentativas': vaca['total_tentativas'],
        'total_prenhezes': vaca['total_prenhezes'],
        'total_abortos': vaca['total_abortos'],
        'num_estacoes': vaca['num_estacoes'],
        'taxa_global': vaca['taxa_global'],
        'peso_medio_global': vaca['peso_medio_global'],
        'peso_medio_machos': vaca['peso_medio_machos'],
        'peso_medio_femeas': vaca['peso_medio_femeas'],
        'qtd_machos': vaca['qtd_machos'],
        'qtd_femeas': vaca['qtd_femeas'],
        'porcentagem_prot_prenhez': vaca['porcentagem_prot_prenhez']
    }
    # Adiciona taxas de PROT individuais
    for prot in vaca['taxas_prot_individual']:
        taxa, prenhez, total = vaca['taxas_prot_individual'][prot]
        row[f'taxa_{prot}'] = taxa
        row[f'prenhez_{prot}'] = prenhez
        row[f'total_{prot}'] = total
    vacas_rows.append(row)
df_vacas = pd.DataFrame(vacas_rows)

# Página 2: estações
estacoes_rows = []
for est in estacoes_data:
    row = {
        'estacao': est['estacao'],
        'total_tentativas': est['total_tentativas'],
        'total_concepcoes': est['total_concepcoes'],
        'taxa_global': est['taxa_global'],
        'data_ia_media': est['data_ia_media']
    }
    # Taxas por categoria
    for cat, (taxa_cat, concepcoes_cat, total_cat) in est['taxas_categoria'].items():
        row[f'taxa_cat_{cat}'] = taxa_cat
        row[f'conceps_cat_{cat}'] = concepcoes_cat
        row[f'total_cat_{cat}'] = total_cat
    # Taxas por protocolo
    for prot, (taxa, prenhez, total) in est['taxas_prot'].items():
        row[f'taxa_{prot}'] = taxa
        row[f'prenhez_{prot}'] = prenhez
        row[f'total_{prot}'] = total
    estacoes_rows.append(row)
df_estacoes = pd.DataFrame(estacoes_rows)

with pd.ExcelWriter('saida_analise.xlsx', engine='xlsxwriter') as writer:
    df_vacas.to_excel(writer, sheet_name='vacas', index=False)
    df_estacoes.to_excel(writer, sheet_name='estacoes', index=False)

