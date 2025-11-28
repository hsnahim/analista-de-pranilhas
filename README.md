# Analisador de Dados de Estação de Monta

## Descrição

Este script em Python foi desenvolvido para automatizar a análise de dados zootécnicos provenientes de planilhas de controle de estação de monta. A ferramenta processa os dados de um arquivo Excel (`CONTROLE ESTACAO.xlsx`), realiza uma série de cálculos complexos e gera um relatório final detalhado em um novo arquivo Excel, além de gráficos para visualização de tendências.

O objetivo é fornecer uma visão completa do desempenho reprodutivo e produtivo do rebanho, tanto a nível geral (por estação), quanto a nível específico (por categoria de animal e por protocolo de IATF).

## Funcionalidades Principais

- **Análise Multi-aba**: Processa múltiplas abas (geralmente representando diferentes anos ou estações) de um arquivo Excel.
- **Seleção Interativa**: Pergunta ao usuário quantas das abas mais recentes ele deseja analisar, tornando a análise flexível.
- **Relatório Multi-página em Excel**: Gera um arquivo `saida_analise.xlsx` com 4 abas distintas:
    1.  `Vacas`: Análise consolidada do histórico de vida de cada animal.
    2.  `Estacoes_Global`: Resumo com os totais e médias de cada estação analisada.
    3.  `Estacoes_por_Categoria`: Análise detalhada do desempenho de cada categoria (Novilha, Vaca Solteira, etc.) dentro de cada estação.
    4.  `Estacoes_por_Protocolo`: Análise do desempenho de cada protocolo de IATF para a estação inteira.
- **Geração Automática de Gráficos**: Cria gráficos de barras para os principais indicadores e os salva em uma pasta `graficos`, organizados em subpastas (`geral` e `categorias`).
- **Lógica de Negócio Customizada**: Interpreta regras complexas da coluna `HISTÓRICO` para calcular participações, concepções e abortos por protocolo.
- **Cálculos Detalhados**: Inclui métricas como:
    - Taxas de prenhez (geral, por categoria, por protocolo, por animal).
    - Contagem de abortos.
    - Peso médio de desmame (`PESO 205`) por sexo.
    - Data média de IA.

## Pré-requisitos

-   Python 3.7 ou superior
-   O arquivo `PLANILHA ATUAL.xlsx` devidamente formatado na mesma pasta do projeto.

## Instalação

É altamente recomendado usar um ambiente virtual (`.venv`) para manter as dependências do projeto isoladas.

1.  **Clone ou baixe este projeto.**

2.  **Crie e ative o ambiente virtual:**
    Abra um terminal na pasta do projeto e execute:
    ```bash
    # Cria a pasta .venv
    python -m venv .venv
    ```
    Para ativar o ambiente:
    -   No **Windows**:
        ```powershell
        .\.venv\Scripts\activate
        ```
    -   No **Linux ou macOS**:
        ```bash
        source .venv/bin/activate
        ```

3.  **Instale as dependências:**
    Com o ambiente ativado, instale todas as bibliotecas necessárias de uma só vez usando o arquivo `requirements.txt`:
    ```bash
    pip install -r requirements.txt
    ```

## Estrutura do Arquivo de Entrada

O script espera um arquivo chamado `CONTROLE ESTACAO.xlsx` na mesma pasta. As abas devem estar em ordem cronológica (as mais antigas primeiro, as mais recentes por último).

As planilhas devem conter as seguintes colunas (os nomes não precisam ser exatos, mas devem ser reconhecíveis):

-   `ANIMAL`: Identificação única do animal.
-   `HISTÓRICO`: Campo de texto com os protocolos aplicados.
-   `SITUAÇÃO`: Resultado da inseminação (P, R, AB, etc.).
-   `CATEGORIA`: Categoria do animal na estação (Novilha, Primípara, etc.).
-   `DATA IA`: Data da inseminação.
-   `PESO 205`: Peso ajustado aos 205 dias do bezerro.
-   `SEXO`: Sexo do bezerro (M ou F).

## Como Usar

1.  Certifique-se de que seu ambiente virtual esteja **ativado**.
2.  Coloque o arquivo `CONTROLE ESTACAO.xlsx` na mesma pasta que o script `analize_planilha.py`.
3.  Execute o script através do terminal:
    ```bash
    python analize_planilha.py
    ```
4.  O programa irá listar as abas encontradas e perguntará quantas das mais recentes você deseja analisar.
    -   Digite um número (ex: `3` para analisar as 3 últimas) e pressione Enter.
    -   Ou apenas pressione Enter para analisar todas as abas.
5.  Aguarde a execução. O script irá imprimir mensagens de progresso no terminal.
6.  Ao final, os arquivos de saída serão gerados na pasta.

## Arquivos de Saída

-   **`saida_analise.xlsx`**: O relatório principal em formato Excel, contendo as 4 abas de análise.
-   **Pasta `graficos`**:
    -   **Subpasta `geral`**: Contém os gráficos de desempenho geral ao longo das estações.
    -   **Subpasta `categorias`**: Contém os gráficos de desempenho por categoria, gerados para cada estação analisada.
