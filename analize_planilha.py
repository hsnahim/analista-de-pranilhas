# --- Bloco de Importação de Bibliotecas ---
# pandas: Essencial para manipulação de dados, como ler e analisar a planilha Excel.
import pandas as pd
# matplotlib.pyplot: Usado para criar os gráficos de barras.
import matplotlib.pyplot as plt
# re: Módulo de expressões regulares, usado para fazer buscas complexas nos textos do histórico.
import re
# os: Módulo para interagir com o sistema operacional, usado aqui para criar pastas para os gráficos.
import os

# --- Carregamento do Arquivo Excel ---
# Tenta carregar o arquivo. Se não encontrar, exibe uma mensagem de erro amigável e encerra o programa.
try:
    xl_file = pd.ExcelFile('CONTROLE ESTACAO.xlsx')
except FileNotFoundError:
    print("ERRO: O arquivo 'CONTROLE ESTACAO.xlsx' não foi encontrado. Verifique se o nome está correto e se ele está na mesma pasta do script.")
    exit()


# --- Seção de Interação com o Usuário ---
# Esta seção permite que o usuário escolha dinamicamente quantas abas (anos) recentes ele quer analisar.
abas_disponiveis = xl_file.sheet_names
print(f"Abas encontradas na planilha: {abas_disponiveis}")

# Pede ao usuário para digitar um número.
num_abas_a_usar = input(f"\nDigite o número de abas mais recentes a serem analisadas (ex: 2 para usar as duas últimas). \nPressione Enter para analisar todas as {len(abas_disponiveis)} abas: ")

abas_selecionadas = []
try:
    # Se o usuário apertar Enter, a entrada será vazia.
    if num_abas_a_usar.strip() == "":
        abas_selecionadas = abas_disponiveis
        print(f"Analisando todas as {len(abas_selecionadas)} abas.")
    else:
        # Tenta converter a entrada para um número inteiro.
        num = int(num_abas_a_usar)
        if num > 0:
            # Usa a notação de fatiamento de lista `[-num:]` para pegar os últimos `num` itens.
            abas_selecionadas = abas_disponiveis[-num:]
            print(f"Analisando as últimas {len(abas_selecionadas)} abas: {abas_selecionadas}")
        else:
            # Se o número for zero ou negativo, usa todas as abas por segurança.
            abas_selecionadas = abas_disponiveis
            print("Número inválido. Analisando todas as abas por padrão.")
# Se a conversão para inteiro falhar (ex: usuário digitou texto), usa todas as abas.
except ValueError:
    abas_selecionadas = abas_disponiveis
    print("Entrada inválida. Analisando todas as abas por padrão.")

# --- Funções Utilitárias ---
def get_col_indices(header_row, nomes):
    """
    Função auxiliar para encontrar os números (índices) das colunas que precisamos.
    Ela recebe a primeira linha da planilha (cabeçalho) e uma lista de nomes de colunas
    (ex: ['ANIMAL', 'HISTÓRICO']) e retorna um dicionário mapeando o nome ao seu índice.
    Isso torna o código resistente a mudanças na ordem das colunas na planilha.
    """
    indices = {nome: None for nome in nomes}
    for idx, col_name in enumerate(header_row):
        nome_col = str(col_name).strip().upper() # Padroniza o nome da coluna para maiúsculas e sem espaços
        if nome_col in indices:
            indices[nome_col] = idx
    return indices

# --- Leitura e Carregamento dos Dados ---
# Cria um dicionário para armazenar os dados de cada aba selecionada.
# Ler todas as abas de uma vez no início melhora a performance, pois não recarregamos o arquivo repetidamente.
dfs = {}
abas_com_erro = []
# O laço agora itera apenas sobre a lista de abas que o usuário selecionou.
for nome_aba in abas_selecionadas:
    try:
        df = xl_file.parse(nome_aba)
        dfs[nome_aba] = df
    except Exception as e:
        print(f"[AVISO] Erro ao ler a aba '{nome_aba}': {e}")
        abas_com_erro.append(nome_aba)
if abas_com_erro:
    print(f"\nAs seguintes abas não puderam ser lidas e foram ignoradas: {abas_com_erro}")

# --- Geração da Blacklist de Históricos ---
# Cria uma lista (set) de todos os valores da coluna 'HISTÓRICO' que não são considerados válidos
# (não contêm 'PROT' ou 'IA-'). Isso é feito uma vez para otimizar a performance.
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

# --- Análise por Animal (Vaca) ---
# O objetivo desta seção é criar a aba "Vacas", consolidando o histórico de vida de cada animal.
if abas_selecionadas:
    # A lista de animais a serem analisados é retirada da última aba selecionada.
    ultima_aba = abas_selecionadas[-1]
    df_full = dfs[ultima_aba]
    header_row = df_full.iloc[0]
    indices = get_col_indices(header_row, ['ANIMAL'])
    animal_idx = indices['ANIMAL']
    df = df_full.iloc[1:].reset_index(drop=True)
    # Pega apenas os IDs únicos de animais para evitar analisar o mesmo animal múltiplas vezes.
    animal_ids = df.iloc[:, animal_idx].astype(str).str.strip().unique().tolist()
    animal_ids = [aid for aid in animal_ids if aid] # Remove IDs vazios
else:
    animal_ids = []

vacas_data = []
num_animais_ignorados = 0
prot_labels = [f"{i}PROT" for i in range(1, 32)] # Cria a lista de '1PROT' a '31PROT'

# Itera sobre cada animal encontrado.
for animal_id in animal_ids:
    # Primeiro, verifica se o animal tem algum histórico inválido em qualquer uma das abas.
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
    
    # Se for inválido, pula para o próximo animal.
    if animal_invalido:
        num_animais_ignorados += 1
        continue

    # Variáveis para acumular os dados do animal ao longo de todas as abas.
    pesos_machos, pesos_femeas = [], []
    qtd_machos, qtd_femeas = 0, 0
    total_tentativas, total_prenhezes, total_abortos = 0, 0, 0
    prot_total_vaca = {prot: 0 for prot in prot_labels}
    prot_prenhez_vaca = {prot: 0 for prot in prot_labels}
    categoria_recente = "N/A" # Armazena a categoria mais recente da vaca

    # Agora, itera sobre as abas novamente para coletar os dados consolidados.
    for nome_aba, df_full_data in dfs.items():
        header_row_data = df_full_data.iloc[0]
        indices_data = get_col_indices(header_row_data, ['HISTÓRICO', 'SITUAÇÃO', 'PESO 205', 'ANIMAL', 'SEXO', 'CATEGORIA'])
        historico_idx, situacao_idx, peso_idx, animal_idx, sexo_idx, categoria_idx = (
            indices_data['HISTÓRICO'], indices_data['SITUAÇÃO'], indices_data['PESO 205'], 
            indices_data['ANIMAL'], indices_data['SEXO'], indices_data['CATEGORIA']
        )
        if animal_idx is None: continue
        
        df_data = df_full_data.iloc[1:]
        mask_animal = df_data.iloc[:, animal_idx].astype(str).str.strip() == animal_id
        df_animal = df_data[mask_animal]
        if df_animal.empty: continue

        # Captura a categoria da vaca nesta aba. Como o loop está em ordem cronológica,
        # a última categoria encontrada será a mais recente.
        if categoria_idx is not None:
            categorias_encontradas = df_animal.iloc[:, categoria_idx].dropna().astype(str).str.strip().unique()
            if len(categorias_encontradas) > 0:
                categoria_recente = categorias_encontradas[0]

        # Filtra registros com histórico inválido.
        if historico_idx is not None:
            valid_mask = ~df_animal.iloc[:, historico_idx].astype(str).str.strip().isin(blacklist_historicos)
            df_animal = df_animal[valid_mask]
            if df_animal.empty: continue
        
            # Acumula os dados de protocolo para a vaca.
            historicos_vaca = df_animal.iloc[:, historico_idx].astype(str).str.replace(' ', '', regex=False).str.upper()
            for prot in prot_labels:
                prot_total_vaca[prot] += historicos_vaca.str.contains(rf'(?<!\d){prot}(?!\d)', regex=True, na=False).sum()
                prot_prenhez_vaca[prot] += historicos_vaca.str.contains(rf'(?<!\d){prot}-(?:P|AB)', regex=True, na=False).sum()

        # Acumula os dados de peso e sexo dos bezerros.
        if peso_idx is not None and sexo_idx is not None:
            pesos = pd.to_numeric(df_animal.iloc[:, peso_idx], errors='coerce')
            sexos = df_animal.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
            pesos_machos.extend(pesos[sexos == 'M'].dropna().tolist())
            pesos_femeas.extend(pesos[sexos == 'F'].dropna().tolist())
            qtd_machos += (sexos == 'M').sum()
            qtd_femeas += (sexos == 'F').sum()

        # Acumula os dados gerais de tentativas e resultados.
        if situacao_idx is not None:
            situacoes = df_animal.iloc[:, situacao_idx].astype(str)
            total_tentativas += len(situacoes)
            total_prenhezes += situacoes.isin(['P', 'AB', 'P2', 'REAB']).sum()
            total_abortos += situacoes.isin(['AB', 'REAB']).sum()
    
    # Calcula as médias e taxas finais para a vaca.
    peso_medio_global = (sum(pesos_machos) + sum(pesos_femeas)) / (len(pesos_machos) + len(pesos_femeas)) if (pesos_machos or pesos_femeas) else None
    taxa_global_num = total_prenhezes / total_tentativas if total_tentativas > 0 else None
    
    # Monta o dicionário com todos os dados da vaca.
    vaca_info = {
        'animal_id': animal_id,
        'categoria_recente': categoria_recente,
        'total_tentativas': total_tentativas, 
        'total_prenhezes': total_prenhezes,
        'total_abortos': total_abortos, 
        'taxa_global': f"{taxa_global_num * 100:.1f}%" if pd.notna(taxa_global_num) else None,
        'peso_medio_global': peso_medio_global, 
        'qtd_machos': qtd_machos, 
        'qtd_femeas': qtd_femeas,
    }

    # Adiciona os dados de cada protocolo ao dicionário da vaca.
    for prot in prot_labels:
        total = prot_total_vaca[prot]
        prenhez = prot_prenhez_vaca[prot]
        taxa_vaca_num = prenhez / total if total > 0 else None
        vaca_info[f'{prot}_total_vaca'] = total
        vaca_info[f'{prot}_prenhez_vaca'] = prenhez
        vaca_info[f'{prot}_taxa_vaca'] = f"{taxa_vaca_num * 100:.1f}%" if pd.notna(taxa_vaca_num) else None

    vacas_data.append(vaca_info)

# --- Coleta de Dados para as Múltiplas Abas de Estação ---
# Listas para armazenar os dados de cada uma das futuras abas do Excel.
estacoes_global_data = []
analise_categoria_data = []
estacoes_protocolo_data = []

# Itera sobre cada estação (aba) selecionada pelo usuário.
for nome_aba, df_full in dfs.items():
    header_row = df_full.iloc[0]
    indices = get_col_indices(header_row, ['ANIMAL', 'PESO 205', 'SEXO', 'SITUAÇÃO', 'HISTÓRICO', 'CATEGORIA', 'DATA IA'])
    animal_idx, peso_idx, sexo_idx, situacao_idx, historico_idx, categoria_idx, data_ia_idx = (
        indices['ANIMAL'], indices['PESO 205'], indices['SEXO'], indices['SITUAÇÃO'], 
        indices['HISTÓRICO'], indices['CATEGORIA'], indices['DATA IA']
    )
    
    # Pula a aba se alguma coluna essencial não for encontrada.
    if historico_idx is None or situacao_idx is None or categoria_idx is None:
        print(f"[AVISO] Aba '{nome_aba}' ignorada na análise detalhada por falta de colunas essenciais.")
        continue
        
    df = df_full.iloc[1:].reset_index(drop=True)
    # Aplica o filtro da blacklist na estação inteira.
    df = df[~df.iloc[:, historico_idx].astype(str).str.strip().isin(blacklist_historicos)]

    # --- 1. Cálculo para a aba "Estacoes_Global" ---
    # Calcula os totais e médias para a estação inteira.
    situacoes_global = df.iloc[:, situacao_idx].astype(str)
    taxa_prenhez_geral_num = situacoes_global.isin(['P', 'AB', 'P2', 'REAB']).sum() / len(df) if len(df) > 0 else None
    
    peso_medio_machos_global, peso_medio_femeas_global = None, None
    data_media_ia_geral, data_media_ia_machos, data_media_ia_femeas = None, None, None
    if peso_idx is not None and sexo_idx is not None:
        pesos = pd.to_numeric(df.iloc[:, peso_idx], errors='coerce')
        sexos = df.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
        peso_medio_machos_global = pesos[sexos == 'M'].mean()
        peso_medio_femeas_global = pesos[sexos == 'F'].mean()
    if data_ia_idx is not None and sexo_idx is not None:
        datas_ia = pd.to_datetime(df.iloc[:, data_ia_idx], errors='coerce')
        sexos_datas = df.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
        data_media_geral = datas_ia.dropna().mean()
        data_media_ia_machos = datas_ia[sexos_datas == 'M'].dropna().mean()
        data_media_ia_femeas = datas_ia[sexos_datas == 'F'].dropna().mean()
        
    estacoes_global_data.append({
        'estacao': nome_aba, 'total_registros': len(df),
        'total_concepcoes': situacoes_global.isin(['P', 'AB', 'P2', 'REAB']).sum(),
        'total_abortos': situacoes_global.isin(['AB', 'REAB']).sum(),
        'taxa_prenhez_geral': f"{taxa_prenhez_geral_num * 100:.1f}%" if pd.notna(taxa_prenhez_geral_num) else None,
        'peso_medio_machos_geral': peso_medio_machos_global,
        'peso_medio_femeas_geral': peso_medio_femeas_global,
        'data_media_ia_geral': data_media_geral.strftime('%Y-%m-%d') if pd.notna(data_media_geral) else None,
        'data_media_ia_machos': data_media_ia_machos.strftime('%Y-%m-%d') if pd.notna(data_media_ia_machos) else None,
        'data_media_ia_femeas': data_media_ia_femeas.strftime('%Y-%m-%d') if pd.notna(data_media_ia_femeas) else None,
    })
    
    # --- 2. Cálculo para a aba "Estacoes_por_Protocolo" ---
    # Calcula os dados de cada protocolo para a estação inteira.
    prot_stats_estacao = {}
    historicos_estacao = df.iloc[:, historico_idx].astype(str).str.replace(' ', '', regex=False).str.upper()
    for prot in prot_labels:
        mask_total = historicos_estacao.str.contains(rf'(?<!\d){prot}(?!\d)', regex=True, na=False)
        total_prot = mask_total.sum()
        if total_prot > 0:
            mask_prenhez = historicos_estacao.str.contains(rf'(?<!\d){prot}-(?:P|AB)', regex=True, na=False)
            mask_contem_aborto = historicos_estacao.str.contains(r'-AB', na=False)
            mask_aborto_final = mask_total & mask_contem_aborto
            taxa_prot_num = mask_prenhez.sum() / total_prot
            prot_stats_estacao[prot] = {
                'total': total_prot, 'prenhezes': mask_prenhez.sum(),
                'taxa': f"{taxa_prot_num * 100:.1f}%" if pd.notna(taxa_prot_num) else None,
                'abortos': mask_aborto_final.sum()
            }
    estacoes_protocolo_data.append({'estacao': nome_aba, 'prots': prot_stats_estacao})
    
    # --- 3. Cálculo para a aba "Estacoes_por_Categoria" ---
    # Itera sobre cada categoria única da estação para calcular seus dados específicos.
    categorias = df.iloc[:, categoria_idx].dropna().astype(str).str.strip().unique()
    for cat in categorias:
        if not cat: continue
        df_cat = df[df.iloc[:, categoria_idx].astype(str).str.strip() == cat]
        situacoes_cat = df_cat.iloc[:, situacao_idx].astype(str)
        taxa_prenhez_cat_num = situacoes_cat.isin(['P', 'AB', 'P2', 'REAB']).sum() / len(df_cat) if len(df_cat) > 0 else None
        
        peso_medio_machos_cat, peso_medio_femeas_cat = None, None
        data_media_ia_geral_cat, data_media_ia_machos_cat, data_media_ia_femeas_cat = None, None, None
        if peso_idx is not None and sexo_idx is not None:
            pesos_cat = pd.to_numeric(df_cat.iloc[:, peso_idx], errors='coerce')
            sexos_cat = df_cat.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
            peso_medio_machos_cat = pesos_cat[sexos_cat == 'M'].mean()
            peso_medio_femeas_cat = pesos_cat[sexos_cat == 'F'].mean()
        if data_ia_idx is not None and sexo_idx is not None:
            datas_ia_cat = pd.to_datetime(df_cat.iloc[:, data_ia_idx], errors='coerce')
            sexos_datas_cat = df_cat.iloc[:, sexo_idx].astype(str).str.upper().str.strip()
            data_media_geral_cat = datas_ia_cat.dropna().mean()
            data_media_ia_machos_cat = datas_ia_cat[sexos_datas_cat == 'M'].dropna().mean()
            data_media_ia_femeas_cat = datas_ia_cat[sexos_datas_cat == 'F'].dropna().mean()
        analise_categoria_data.append({
            'estacao': nome_aba, 'categoria': cat, 'total_registros_cat': len(df_cat),
            'total_concepcoes_cat': situacoes_cat.isin(['P', 'AB', 'P2', 'REAB']).sum(),
            'total_abortos_cat': situacoes_cat.isin(['AB', 'REAB']).sum(),
            'taxa_prenhez_cat': f"{taxa_prenhez_cat_num * 100:.1f}%" if pd.notna(taxa_prenhez_cat_num) else None,
            'peso_medio_machos_cat': peso_medio_machos_cat,
            'peso_medio_femeas_cat': peso_medio_femeas_cat,
            'data_media_ia_geral_cat': data_media_geral_cat.strftime('%Y-%m-%d') if pd.notna(data_media_geral_cat) else None,
            'data_media_ia_machos_cat': data_media_ia_machos_cat.strftime('%Y-%m-%d') if pd.notna(data_media_ia_machos_cat) else None,
            'data_media_ia_femeas_cat': data_media_ia_femeas_cat.strftime('%Y-%m-%d') if pd.notna(data_media_ia_femeas_cat) else None,
        })

# --- Geração dos DataFrames Finais ---
# Converte as listas de dicionários em DataFrames do Pandas.
def expand_stats_detalhado(data, prot_labels):
    """
    Função auxiliar para transformar os dados de protocolo aninhados em colunas.
    Ex: {'prots': {'1PROT': {'total': 5}}} -> colunas '1PROT_total', '1PROT_prenhezes', etc.
    """
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
        expanded.append(new_row)
    return expanded

df_vacas = pd.DataFrame(vacas_data)
# Reordena as colunas da aba Vacas para colocar a categoria logo após o ID.
if not df_vacas.empty:
    cols_primeiras = ['animal_id', 'categoria_recente']
    cols_restantes = [col for col in df_vacas.columns if col not in cols_primeiras]
    df_vacas = df_vacas[cols_primeiras + cols_restantes]

df_estacoes_global = pd.DataFrame(estacoes_global_data)
df_estacoes_categoria = pd.DataFrame(analise_categoria_data)
df_estacoes_protocolo = pd.DataFrame(expand_stats_detalhado(estacoes_protocolo_data, prot_labels))

# --- Geração do Arquivo Excel (método padrão) ---
# Usa o ExcelWriter para salvar múltiplos DataFrames em diferentes abas no mesmo arquivo.
with pd.ExcelWriter('saida_analise.xlsx') as writer:
    df_vacas.to_excel(writer, sheet_name='Vacas', index=False)
    df_estacoes_global.to_excel(writer, sheet_name='Estacoes_Global', index=False)
    df_estacoes_categoria.to_excel(writer, sheet_name='Estacoes_por_Categoria', index=False)
    df_estacoes_protocolo.to_excel(writer, sheet_name='Estacoes_por_Protocolo', index=False)

print("\nArquivo 'saida_analise.xlsx' gerado com sucesso!")
print(f"A análise considerou as abas: {abas_selecionadas}")

# --- Seção de Geração de Gráficos ---
if not df_estacoes_global.empty:
    # Cria a pasta principal 'graficos' e as subpastas 'geral' e 'categorias'.
    output_dir_base = "graficos"
    output_dir_geral = os.path.join(output_dir_base, "geral")
    output_dir_categorias = os.path.join(output_dir_base, "categorias")

    if not os.path.exists(output_dir_geral):
        os.makedirs(output_dir_geral)
    if not os.path.exists(output_dir_categorias):
        os.makedirs(output_dir_categorias)
    print(f"Gráficos serão salvos nas subpastas dentro de: '{output_dir_base}'")

    def criar_grafico_barras(x_data, y_data, titulo, eixo_y_label, caminho_arquivo, is_percent=False):
        """
        Função auxiliar para criar e salvar um gráfico de barras.
        Ela padroniza a aparência de todos os gráficos.
        """
        plt.figure(figsize=(12, 7)) # Define o tamanho da imagem
        
        # Como as taxas agora são texto (ex: "55,0%"), precisamos convertê-las de volta para número para o gráfico.
        y_data_numeric = y_data
        if is_percent:
            y_data_numeric = pd.to_numeric(y_data.str.replace('%', '', regex=False).str.replace(',', '.', regex=False), errors='coerce') / 100

        bars = plt.bar(x_data.astype(str), y_data_numeric, color='skyblue')
        plt.title(titulo, fontsize=16)
        plt.ylabel(eixo_y_label, fontsize=12)
        plt.xticks(rotation=45, ha='right') # Rotaciona os labels do eixo X para não sobreporem
        
        # Formata o eixo Y como porcentagem, se aplicável
        if is_percent:
            from matplotlib.ticker import PercentFormatter
            plt.gca().yaxis.set_major_formatter(PercentFormatter(1))

        # Adiciona o valor exato (como texto) no topo de cada barra para clareza.
        for i, bar in enumerate(bars):
            label = y_data.iloc[i]
            if pd.isna(label): continue
            plt.text(bar.get_x() + bar.get_width()/2.0, bar.get_height(), label, va='bottom', ha='center')
        
        plt.tight_layout() # Ajusta o layout para garantir que tudo caiba na imagem
        plt.savefig(caminho_arquivo) # Salva o gráfico no arquivo especificado
        plt.close() # Fecha a figura para liberar memória, essencial ao gerar muitos gráficos.

    # 1. Gera os gráficos para a aba "Estacoes_Global".
    df_global_sorted = df_estacoes_global.sort_values('estacao')
    graficos_globais = [
        ('total_concepcoes', 'Total de Concepções por Estação', 'Nº de Concepções'),
        ('total_registros', 'Total de Registros por Estação', 'Nº de Registros'),
        ('total_abortos', 'Total de Abortos por Estação', 'Nº de Abortos'),
        ('taxa_prenhez_geral', 'Taxa de Prenhez Geral por Estação', 'Taxa de Prenhez')
    ]
    for coluna, titulo, eixo_y in graficos_globais:
        caminho = os.path.join(output_dir_geral, f"geral_{coluna}.png")
        is_percent = 'taxa' in coluna
        df_plot = df_global_sorted.dropna(subset=[coluna])
        if not df_plot.empty:
            criar_grafico_barras(df_plot['estacao'], df_plot[coluna], titulo, eixo_y, caminho, is_percent=is_percent)
    print("Gráficos gerais por estação gerados com sucesso.")

    # 2. Gera os gráficos para a aba "Estacoes_por_Categoria", um conjunto para cada ano.
    anos = df_estacoes_categoria['estacao'].unique()
    graficos_categoria = [
        ('total_concepcoes_cat', 'Total de Concepções por Categoria', 'Nº de Concepções'),
        ('total_registros_cat', 'Total de Registros por Categoria', 'Nº de Registros'),
        ('total_abortos_cat', 'Total de Abortos por Categoria', 'Nº de Abortos'),
        ('taxa_prenhez_cat', 'Taxa de Prenhez por Categoria', 'Taxa de Prenhez')
    ]
    for ano in anos:
        df_ano = df_estacoes_categoria[df_estacoes_categoria['estacao'] == ano].sort_values('categoria')
        for coluna, titulo_base, eixo_y in graficos_categoria:
            titulo_completo = f'{titulo_base} - {ano}'
            caminho = os.path.join(output_dir_categorias, f"categoria_{ano}_{coluna}.png")
            is_percent = 'taxa' in coluna
            df_plot = df_ano.dropna(subset=[coluna])
            if not df_plot.empty:
                criar_grafico_barras(df_plot['categoria'], df_plot[coluna], titulo_completo, eixo_y, caminho, is_percent=is_percent)
    print("Gráficos por categoria para cada estação gerados com sucesso.")

# --- Exibição de informações no console ---
# Imprime um resumo final para o usuário no terminal.
print("\nHistóricos inválidos encontrados (ignorados na análise):")
if blacklist_historicos:
    for h in sorted(list(blacklist_historicos)):
        print(f"  - {h}")
else:
    print("  Nenhum histórico inválido encontrado.")
print(f"\nTotal de vacas ignoradas na análise individual (por terem histórico inválido em algum momento): {num_animais_ignorados}")