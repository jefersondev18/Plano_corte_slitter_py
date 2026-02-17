**PLANO DE CORTE**

Guia Completo do Script Python

*Versão 2 --- com Limite de Cortes e Cálculo de KG*

**1. O Que o Script Faz --- Visão Geral**

O script resolve um problema clássico da indústria de tubos e perfis:
**como aproveitar ao máximo a largura de uma bobina de aço**, combinando
cortes de diferentes perfis (matrizes) de forma que a perda de material
fique dentro de uma faixa aceitável.

O usuário informa um perfil principal chamado **âncora** ---
obrigatoriamente produzido --- e o script descobre quais outros perfis
(**complementares**) preenchem o espaço restante da bobina dentro dos
limites de perda.

+-----------------------------------------------------------------------+
| **💡 Analogia**                                                       |
|                                                                       |
| Pense em uma prateleira de 1200 cm. Você precisa colocar caixas de    |
| tamanhos variados. O objetivo é que a soma das caixas deixe uma folga |
| entre 0,67% e 1,70% do espaço total --- nem pouco nem muito.          |
+-----------------------------------------------------------------------+

**1.1 Conceitos fundamentais**

  ---------------------- ------------------------------------------------
  **Termo**              **Significado**

  Matriz                 Perfil de tubo ou barra (ex: 50,80-2\", 100X50).
                         Define o Desenvolvimento.

  Desenvolvimento (mm)   Largura da tira de aço necessária para fabricar
                         aquele perfil.

  Largura da Bobina      Largura fixa do aço: 1000, 1200 ou 1500 mm
                         (padrões industriais).

  Âncora                 O perfil que o usuário quer produzir ---
                         OBRIGATÓRIO, com N ≥ 1 corte.

  Complementar           Outros perfis que ocupam o espaço restante da
                         bobina.

  Combinação             Âncora(N) + Comp1(N1) + Comp2(N2) --- um plano
                         de corte completo.

  Total de cortes        Soma de todos os N da combinação: N_âncora +
                         N_comp1 + N_comp2.

  Perda (%)              (Largura − Soma dos cortes) / Largura × 100.
                         Deve ficar entre 0,67% e 1,70%.
  ---------------------- ------------------------------------------------

**2. Parâmetros de Negócio**

No topo do script ficam as constantes que controlam todas as regras. São
os únicos valores a alterar se as regras mudarem:

+-----------------------------------------------------------------------+
| LARGURAS_BOBINA = \[1200, 1000, 1500\] \# ordem de tentativa          |
|                                                                       |
| PERDA_MIN_PCT = 0.67 \# % mínimo de perda aceito                      |
|                                                                       |
| PERDA_MAX_PCT = 1.70 \# % máximo de perda aceito                      |
|                                                                       |
| MAX_COMP_NA_COMBO = 2 \# máx de complementares por combinação         |
|                                                                       |
| PESO_MEDIO_BOB_PAD = 12_000 \# kg padrão (12 ton) se não informado    |
|                                                                       |
| QTD_BOBINAS_PAD = 1 \# qtd padrão de bobinas se não informado         |
+-----------------------------------------------------------------------+

  --------------------- --------------------------- ----------------------
  **Constante**         **O que controla**          **Exemplo de ajuste**

  LARGURAS_BOBINA       Quais larguras testar e em  Adicionar 1800:
                        que ordem                   \[\..., 1800\]

  PERDA_MIN/MAX_PCT     A janela de perda aceitável Ampliar para 0.50% --
                                                    2.00%

  MAX_COMP_NA_COMBO     Qtd. de complementares por  Mudar para 3 --- mais
                        combinação (afeta           lento
                        performance)                

  PESO_MEDIO_BOB_PAD    Peso padrão quando o        15000 para bobinas
                        usuário deixa em branco     mais pesadas

  QTD_BOBINAS_PAD       Quantidade padrão quando o  Manter em 1
                        usuário deixa em branco     
  --------------------- --------------------------- ----------------------

**3. Fluxo Completo de Execução**

Ao rodar o script, estes são os passos executados em sequência:

  --------------- ----------------------------------------------------------
  **Etapa**       **O que acontece**

  1\. Carga       carregar_dados() lê o Excel, limpa os dados e valida os
                  campos essenciais.

  2\. Menu \[1\]  Usuário escolhe a ESPESSURA entre as disponíveis na base.

  3\. Menu \[2\]  Usuário escolhe o TIPO DE MATERIAL filtrado pela
                  espessura.

  4\. Menu \[3\]  Usuário escolhe a MATRIZ ÂNCORA filtrada por espessura +
                  tipo.

  5\. Menu \[4\]  Usuário informa o LIMITE DE CORTES (opcional --- pode
                  deixar em branco).

  6\. Menu \[5\]  Usuário informa a QUANTIDADE DE BOBINAS (padrão: 1).

  7\. Menu \[6\]  Usuário informa o PESO do lote de bobinas em kg (padrão:
                  12.000 kg).

  8\. Busca       encontrar_combinacoes() tenta 1200 → 1000 → 1500 mm e para
                  no primeiro com resultado.

  9\.             \_buscar_para_largura() avalia todas as combinações
  Combinatorial   válidas para a largura escolhida.

  10\. Exibição   exibir() imprime o resumo no terminal com os parâmetros
                  usados.

  11\. Exportação exportar_xlsx() gera o .xlsx com as duas abas de
                  resultado.
  --------------- ----------------------------------------------------------

**4. A Interface de Usuário --- Os 6 Passos do Menu**

Ao executar o script, o terminal apresenta 6 perguntas em sequência.
Cada uma filtra as opções da anterior, evitando escolhas inválidas.

**Passos \[1\], \[2\] e \[3\] --- Espessura, Tipo e Âncora**

O script exibe apenas as opções que existem na base para os parâmetros
já escolhidos. O usuário nunca verá uma combinação impossível. O Passo
\[3\] mostra o **Desenvolvimento (mm)** de cada matriz para auxiliar na
escolha da âncora.

**Passo \[4\] --- Limite de Cortes ⚙**

**Esta é a restrição de máquina.** Se a máquina suporta no máximo N
cortes simultâneos por bobina, qualquer combinação cuja **soma total de
cortes** ultrapasse esse número é automaticamente descartada.

+-----------------------------------------------------------------------+
| \[4\] Limite máximo de cortes por combinação (restrição de máquina)   |
|                                                                       |
| Deixe em branco e pressione Enter para sem limite.                    |
|                                                                       |
| Limite de cortes: 5                                                   |
|                                                                       |
| \# Como o filtro funciona:                                            |
|                                                                       |
| \# âncora(x3) + comp(x2) = 5 cortes → APROVADO ✓                      |
|                                                                       |
| \# âncora(x2) + comp(x4) = 6 cortes → REJEITADO ✗                     |
|                                                                       |
| \# âncora(x1) = 1 corte → APROVADO ✓                                  |
+-----------------------------------------------------------------------+

+-----------------------------------------------------------------------+
| **ℹ Sem limite**                                                      |
|                                                                       |
| Se o campo for deixado em branco (Enter), o script buscará todas as   |
| combinações válidas independentemente de quantos cortes totais        |
| tenham.                                                               |
+-----------------------------------------------------------------------+

**Passo \[5\] --- Quantidade de Bobinas ⚙**

Informa quantas bobinas serão processadas nesta rodada. Usado no cálculo
de KG. Se deixado em branco, assume 1 bobina.

**Passo \[6\] --- Peso das Bobinas ⚙**

Informa o **peso total do lote de bobinas em kg**. Se deixado em branco,
assume 12.000 kg (12 ton).

+-----------------------------------------------------------------------+
| **💡 Como o peso médio é calculado**                                  |
|                                                                       |
| **Peso médio por bobina** = Peso informado ÷ Quantidade de bobinas    |
| Exemplo: 48.000 kg informados, 4 bobinas → peso médio = **12.000 kg   |
| por bobina** Esse valor é o que entra diretamente na fórmula de KG    |
| --- garantindo proporcionalidade.                                     |
+-----------------------------------------------------------------------+

**5. A Lógica de Busca**

**5.1 Seleção da largura da bobina**

A largura da bobina não é livre --- ela segue os padrões industriais. O
script tenta as larguras em ordem fixa, parando na primeira que
encontrar combinações válidas:

+-----------------------------------------------------------------------+
| LARGURAS_BOBINA = \[1200, 1000, 1500\]                                |
|                                                                       |
| para largura em \[1200, 1000, 1500\]:                                 |
|                                                                       |
| calcular combinações\...                                              |
|                                                                       |
| se encontrou alguma → parar aqui ✓                                    |
|                                                                       |
| se não encontrou → tentar próxima                                     |
+-----------------------------------------------------------------------+

Começar por **1200 mm** é a escolha mais comum e que oferece mais opções
de combinação. A 1000 mm é mais restritiva e a 1500 mm é usada apenas
como último recurso.

**5.2 Fórmulas de validação**

Para cada combinação testada, o script calcula:

+-----------------------------------------------------------------------+
| Soma dos cortes = Σ (Desenvolvimento_i × N_cortes_i)                  |
|                                                                       |
| Perda (mm) = Largura_bobina − Soma dos cortes                         |
|                                                                       |
| Perda (%) = Perda (mm) / Largura_bobina × 100                         |
|                                                                       |
| Uma combinação é VÁLIDA quando:                                       |
|                                                                       |
| 0,67% ≤ Perda (%) ≤ 1,70%                                             |
|                                                                       |
| E Total_cortes ≤ Limite_cortes (se informado)                         |
+-----------------------------------------------------------------------+

+-----------------------------------------------------------------------+
| **📐 Exemplo numérico**                                               |
|                                                                       |
| **Largura:** 1.200 mm **Âncora:** 50,80-2\" \| dev = 157 mm \| N = 3  |
| → 157 × 3 = 471 mm **Comp 1:** 38,10-1.1/2\" \| dev = 117 mm \| N = 6 |
| → 117 × 6 = 702 mm **Soma:** 471 + 702 = 1.173 mm \| **Perda:** 1.200 |
| − 1.173 = 27 mm (2,25%) → **REJEITADA** **Total de cortes:** 3 + 6 =  |
| 9 → se limite = 5, seria rejeitada por cortes também                  |
+-----------------------------------------------------------------------+

**5.3 O motor combinatorial**

A função \_buscar_para_largura() testa sistematicamente todas as
possibilidades para uma largura:

+-----------------------------------------------------------------------+
| para cada N_âncora de 1 até N_máximo:                                 |
|                                                                       |
| espaco_restante = largura − (dev_âncora × N_âncora)                   |
|                                                                       |
| \# Caso 1: só a âncora                                                |
|                                                                       |
| se perda válida E total_cortes ≤ limite:                              |
|                                                                       |
| guardar resultado                                                     |
|                                                                       |
| \# Caso 2: âncora + até 2 complementares                              |
|                                                                       |
| para cada subconjunto de complementares que cabem:                    |
|                                                                       |
| para cada N_cortes possível de cada complementar:                     |
|                                                                       |
| soma_total = soma_âncora + soma_complementares                        |
|                                                                       |
| total_cortes = N_âncora + ΣN_comp                                     |
|                                                                       |
| se perda válida E total_cortes ≤ limite:                              |
|                                                                       |
| guardar resultado                                                     |
+-----------------------------------------------------------------------+

  --------------------- --------------------------- --------------------------
  **Função Python**     **O que faz**               **Exemplo**

  combinations(lista,   Gera todos os subconjuntos  combinations(\[A,B,C\], 2)
  n)                    de tamanho n sem repetir    → (A,B), (A,C), (B,C)
                        elementos                   

  product(r1, r2, \...) Produto cartesiano ---      product(\[1,2\],\[1,3\]) →
                        todas as combinações de     (1,1),(1,3),(2,1),(2,3)
                        quantidades de cortes       
  --------------------- --------------------------- --------------------------

**6. Cálculo de Quantidade em KG ⚙**

**6.1 A fórmula completa**

O cálculo acontece em três etapas dentro do script:

+-----------------------------------------------------------------------+
| \# Etapa 1 --- peso médio por bobina (calculado automaticamente)      |
|                                                                       |
| peso_medio_calc = peso_informado / qtd_bobinas                        |
|                                                                       |
| \# Etapa 2 --- KG de cada perfil individualmente                      |
|                                                                       |
| KG_i = (peso_medio_calc / largura_bobina) × (N_cortes_i ×             |
| Desenvolvimento_i × qtd_bobinas)                                      |
|                                                                       |
| \# Etapa 3 --- KG total da combinação                                 |
|                                                                       |
| KG_combo = KG_âncora + KG_comp1 + KG_comp2                            |
+-----------------------------------------------------------------------+

**6.2 Exemplos práticos**

  ------------------ ------------ ---------------- -------------------------
  **Peso Informado** **Qtd        **Peso           **Interpretação**
                     Bobinas**    Médio/Bobina**   

  12.000 kg (padrão) 1 (padrão)   12.000 kg        1 bobina de 12 ton ---
                                                   comportamento padrão

  48.000 kg          4            12.000 kg        4 bobinas de 12 ton cada

  60.000 kg          4            15.000 kg        4 bobinas mais pesadas,
                                                   de 15 ton cada

  24.000 kg          4            6.000 kg         4 bobinas leves, de 6 ton
                                                   cada

  15.000 kg          1            15.000 kg        1 bobina de 15 ton
  ------------------ ------------ ---------------- -------------------------

+-----------------------------------------------------------------------+
| **🔎 Por que dividir o peso pela quantidade?**                        |
|                                                                       |
| O usuário informa o peso **total do lote**. Dividir por qtd_bobinas   |
| dá o peso médio de **uma bobina**. Isso garante proporcionalidade: se |
| você tem 4 bobinas de 12 ton, o KG de um corte que ocupa 50% da       |
| largura equivale a 50% do peso de uma bobina, multiplicado pelas 4    |
| bobinas.                                                              |
+-----------------------------------------------------------------------+

**7. Saídas do Script**

**7.1 Terminal**

O terminal exibe um resumo com os parâmetros usados e a lista de
combinações ordenada da menor para a maior perda:

+-----------------------------------------------------------------------+
| ============================================================          |
|                                                                       |
| PLANO DE CORTE --- COMBINAÇÕES VÁLIDAS                                |
|                                                                       |
| ============================================================          |
|                                                                       |
| Âncora : 50,80-2\"                                                    |
|                                                                       |
| Espessura : 2.0 mm                                                    |
|                                                                       |
| Largura bobina : 1200 mm (padrão usado)                               |
|                                                                       |
| Janela de perda: 0.67% -- 1.70% \| 8.04 mm -- 20.40 mm                |
|                                                                       |
| Limite cortes : 5 cortes (soma total por combinação)                  |
|                                                                       |
| Combinações : 104                                                     |
|                                                                       |
| \# Combinação Soma(mm) Perda(mm) Perda(%)                             |
|                                                                       |
| 1 50,80-2\"(x1) + 152,40-6\"(x1) + \... 1191.00 9.000 0.7500%         |
|                                                                       |
| 2 50,80-2\"(x2) + 101,60-4\"(x2) + \... 1191.00 9.000 0.7500%         |
+-----------------------------------------------------------------------+

**7.2 Excel --- Aba \"Combinações\"**

  ------------- --------------------------- ------------------------------
  **Coluna**    **Conteúdo**                **Destaque visual**

  \#            Número sequencial da        ---
                combinação                  

  Combinação    Descrição: Âncora(xN) +     ---
                Comp(xN)\...                

  N Âncora      Cortes da matriz âncora     Amarelo

  Total Cortes  Soma de TODOS os cortes da  ---
  ⚙             combinação                  

  Soma Cortes   Soma dos desenvolvimentos × ---
  (mm)          cortes                      

  Perda (mm)    Espaço não aproveitado na   ---
                bobina                      

  Perda (%)     Perda em percentual da      ---
                largura total               

  Qtd. KG ⚙     Quilos de aço desta         Roxo
                combinação                  

  Status        ✓ Válida para todas as      ---
                linhas exibidas             
  ------------- --------------------------- ------------------------------

**7.3 Excel --- Aba \"Detalhes\"**

Desmembra cada combinação linha a linha, com o KG individual por perfil:

  ----------------- -------------------------------------------------------
  **Coluna**        **Conteúdo**

  \# Combo          Número que liga esta linha à aba Combinações

  Papel             ÂNCORA (fundo amarelo) ou Complementar (fundo branco)

  Matriz            Nome do perfil

  Desenvolvimento   Largura da tira em mm

  N° Cortes         Quantos cortes deste perfil

  Subtotal (mm)     Desenvolvimento × N° Cortes

  Qtd. KG ⚙         KG deste perfil específico (fundo roxo)
  ----------------- -------------------------------------------------------

O cabeçalho do Excel exibe todos os parâmetros da sessão: **Peso
Informado**, **Peso Médio por Bobina**, **Qtd. de Bobinas**, **Limite de
Cortes** e **Largura usada** --- para rastreabilidade completa.

**8. Como Personalizar**

**8.1 Ajustes rápidos --- sem conhecer Python**

  --------------------- ------------------------- -----------------------
  **O que mudar**       **Onde no código**        **Como fazer**

  Janela de perda       PERDA_MIN_PCT /           Trocar os valores
                        PERDA_MAX_PCT             numéricos

  Ordem das larguras    LARGURAS_BOBINA = \[1200, Reordenar ou adicionar
                        \...\]                    valores

  Mais complementares   MAX_COMP_NA_COMBO = 2     Aumentar para 3 (mais
                                                  lento)

  Peso padrão           PESO_MEDIO_BOB_PAD =      Ajustar para o padrão
                        12_000                    da operação

  Qtd. padrão de        QTD_BOBINAS_PAD = 1       Manter em 1 na maioria
  bobinas                                         dos casos
  --------------------- ------------------------- -----------------------

**8.2 Avisos importantes**

-   **Performance:** MAX_COMP_NA_COMBO = 3 pode multiplicar o tempo de
    cálculo por 10× ou mais. Recomendo manter em 2.

-   **Arquivo de entrada:** o script lê sempre db_plano_corte.xlsx da
    pasta BASE_INPUT. Nome e colunas devem ser preservados.

-   **Limite de cortes muito restritivo:** se o limite for muito baixo,
    o resultado pode ser vazio. O terminal avisará sem travar.

-   **Campo de peso:** informe o peso TOTAL do lote, não por bobina. A
    divisão é feita automaticamente.

**9. Glossário**

  ---------------------------------- ----------------------------------------------
  **Termo Python**                   **Explicação em linguagem simples**

  DataFrame                          Tabela em memória, como uma planilha --- com
                                     linhas e colunas filtráveis.

  combinations(lista, n)             Gera todos os subconjuntos de tamanho n. Ex:
                                     (A,B), (A,C), (B,C) de \[A,B,C\].

  product(r1, r2, \...)              Produto cartesiano: todas as combinações entre
                                     sequências.

  int \| None                        Tipo que aceita um inteiro OU None (ausência
                                     de valor). Usado no limite de cortes.

  groupby().mean()                   Agrupa linhas por coluna e calcula a média.
                                     Usado para desenvolvimentos duplicados.

  dropna()                           Remove linhas com valores ausentes (NaN) nas
                                     colunas especificadas.

  astype(str).str.strip()            Converte para texto e remove espaços
                                     invisíveis nas bordas da célula.

  pd.to_numeric(errors=\'coerce\')   Converte para número; se não conseguir, coloca
                                     NaN em vez de travar o programa.
  ---------------------------------- ----------------------------------------------

*Plano de Corte --- Guia v2 \| Gerado com Claude/Anthropic*
