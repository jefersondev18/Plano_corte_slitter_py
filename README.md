# Plano de Corte - Otimizador de Combinações de Matrizes

Sistema inteligente para otimização de corte de bobinas de aço, desenvolvido para a indústria de tubos e perfis metálicos.

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Production-success.svg)]()

---

## Índice

- [Sobre o Projeto](#-sobre-o-projeto)
- [Problema que Resolve](#-problema-que-resolve)
- [Funcionalidades](#-funcionalidades)
- [Base Matemática](#-base-matemática)
- [Instalação](#-instalação)
- [Como Usar](#-como-usar)
- [Estrutura do Código](#-estrutura-do-código)
- [Exemplos](#-exemplos)
- [Parâmetros de Configuração](#️-parâmetros-de-configuração)
- [Regras de Negócio](#-regras-de-negócio)
- [Contribuindo](#-contribuindo)
- [Licença](#-licença)

---

## Sobre o Projeto

O **Plano de Corte** é um sistema desenvolvido para otimizar o aproveitamento de bobinas de aço na fabricação de tubos e perfis metálicos. Ele calcula automaticamente as melhores combinações de corte, minimizando desperdício de material e respeitando restrições de máquina e qualidade.

### Características Principais

- **Otimização Exata**: Testa todas as combinações possíveis dentro das restrições
- **Exportação Excel**: Gera relatórios detalhados em duas abas
- **Interface CLI Intuitiva**: Menu interativo em português
- **Alta Performance**: Processa milhares de combinações em segundos
- **Validação em Cascata**: Aplica múltiplas regras de qualidade
- **Cálculo de KG**: Estima quantidade de aço consumida

---

## 🔍 Problema que Resolve

### Contexto Industrial

Na fabricação de tubos metálicos, bobinas de aço com largura fixa (1000, 1200 ou 1500 mm) são cortadas longitudinalmente para produzir tiras que formarão diferentes perfis. O desafio é:

1. Combinar diferentes perfis (matrizes) na mesma bobina
2. Minimizar o desperdício (refilo/perda)
3. Respeitar limites técnicos de máquina
4. Garantir qualidade mínima de refilo

### Exemplo Prático

**Entrada:**
- Bobina de 1200 mm de largura
- Matriz âncora: `50,80-2"` (desenvolvimento: 157 mm)
- Espessura: 2.0 mm
- Tipo: COMERCIAL

**Saída:**
```
Combinação: 50,80-2"(x1) + 300X75X25(x2) + 21,30-1/2"PT(x1)
Soma: 1190.0 mm
Perda: 10.0 mm (0.83%)
Status: ✓ Válida
```

**Resultado:**
- 1 corte de 50,80-2" (157 mm)
- 2 cortes de 300X75X25 (2 × 394 mm = 788 mm)
- 1 corte de 21,30-1/2"PT (245 mm)
- **Total:** 1190 mm utilizados, 10 mm de refilo (0,83% de perda)

---

## Funcionalidades

### 1. **Busca Combinatorial Inteligente**
- Testa larguras em ordem de preferência (1200 → 1000 → 1500 mm)
- Matriz âncora sempre presente (obrigatória)
- Até 2 matrizes complementares por combinação
- Filtros de viabilidade reduzem espaço de busca

### 2. **Validação Multinível**

| Nível | Validação | Critério |
|-------|-----------|----------|
| 1º | **Perda %** | 0,67% ≤ perda ≤ 1,70% |
| 2º | **Refilo mínimo** | ≥10mm (esp≤3mm) ou ≥14mm (esp>3mm) |
| 3º | **Limite de cortes** | Soma total ≤ limite da máquina |

**Status resultantes:**
- **Válida**: Passou em todas as validações
- **Fora da regra**: Passou na perda % mas não no refilo
- **Rejeitada**: Não atendeu critérios básicos (não aparece)

### 3. **Cálculo de Quantidade (KG)**

Estima o consumo de aço com base em:
- Peso médio das bobinas
- Quantidade de bobinas no lote
- Desenvolvimento de cada matriz
- Número de cortes

**Fórmula:**
```
Peso_médio_bobina = Peso_total_lote / Quantidade_bobinas
KG_matriz = (Peso_médio / Largura) × (N_cortes × Desenvolvimento × Qtd_bobinas)
```

### 4. **Exportação para Excel**

Gera arquivo `.xlsx` com duas abas:

**Aba "Combinações":**
- Cabeçalho com todos os parâmetros usados
- Tabela com todas as combinações válidas
- Colunas: Combinação, N Âncora, Total Cortes, Soma, Perda (mm), Perda (%), KG, Status
- Código de cores: verde (válida), laranja (fora da regra)

**Aba "Detalhes":**
- Detalhamento matriz por matriz
- Identifica âncora vs complementar
- KG individual de cada perfil

---

## Base Matemática

### Modelo do Problema

Este é um **Problema de Empacotamento Unidimensional** (*1D Bin Packing*) com restrições.

**Variáveis:**
- `N_i` = número de cortes da matriz `i` (inteiro ≥ 0)
- `N_ancora ≥ 1` (obrigatório)

**Função Objetivo:**
```
Minimizar: Perda = L - Σ(d_i × N_i)
```

**Restrições:**
```
1. Σ(d_i × N_i) ≤ L                          (capacidade)
2. p_min × L ≤ Perda ≤ p_max × L            (janela de perda)
3. Perda ≥ r_min                             (refilo mínimo)
4. Σ N_i ≤ K                                 (limite de cortes)
5. N_ancora ≥ 1                              (âncora obrigatória)
6. N_i ∈ ℕ                                   (integralidade)
```

### Algoritmo

**Enumeração Completa** com otimizações:

```
Para cada largura L em [1200, 1000, 1500]:
    Para N_ancora = 1 até ⌊L / d_ancora⌋:
        Para cada subconjunto S de complementares:
            Para cada combinação de quantidades (N₁, N₂, ...):
                Se Σ(d_i × N_i) ≤ L e restrições OK:
                    Adicionar aos resultados
```

**Complexidade:** O(L/d × C(n,k) × m^k)
- Típico: ~200.000 combinações testadas em <5 segundos

**Otimizações aplicadas:**
- Ordenação decrescente por desenvolvimento
- Early termination (poda quando ultrapassar largura)
- Filtro prévio de candidatas viáveis
- Limite de profundidade (max 2 complementares)

---

## 🚀 Instalação

### Pré-requisitos

- Python 3.10 ou superior
- pip (gerenciador de pacotes Python)

### Dependências

```bash
pip install pandas openpyxl
```

### Instalação do Projeto

```bash
# Clone o repositório
git clone https://github.com/seu-usuario/plano-corte.git

# Entre no diretório
cd plano-corte

# Instale as dependências
pip install -r requirements.txt
```

### Estrutura de Diretórios

```
plano-corte/
│
├── plano_corte.py          # Script principal
├── requirements.txt         # Dependências Python
├── README.md               # Este arquivo
│
├── files/
│   ├── input/
│   │   └── db_plano_corte.xlsx    # Base de dados de matrizes
│   └── output/                     # Arquivos Excel gerados
│
└── docs/
    └── guia_plano_corte_v2.docx   # Guia completo (opcional)
```

---

## 💻 Como Usar

### Passo 1: Preparar Base de Dados

O arquivo `db_plano_corte.xlsx` deve ter as colunas:

| Coluna | Tipo | Descrição |
|--------|------|-----------|
| Matriz | Texto | Nome do perfil (ex: `50,80-2"`) |
| Espessura | Número | Espessura em mm (ex: `2.0`) |
| Tipo de material | Texto | COMERCIAL, GALVANIZADO, etc |
| Desenvolvimento | Número | Largura necessária em mm (ex: `157.0`) |

### Passo 2: Configurar Caminhos (se necessário)

Edite o topo do arquivo `plano_corte.py`:

```python
# Para Windows
BASE_INPUT  = r'C:\seu\caminho\input'
BASE_OUTPUT = r'C:\seu\caminho\output'

# Para Linux
BASE_INPUT  = r'/home/usuario/plano_corte/input'
BASE_OUTPUT = r'/home/usuario/plano_corte/output'
```

### Passo 3: Executar

```bash
python plano_corte.py
```

### Passo 4: Seguir Menu Interativo

O sistema solicitará 6 informações em sequência:

```
[1] Escolha a espessura: 2.0 mm
[2] Escolha o tipo de material: COMERCIAL
[3] Escolha a matriz âncora: 50,80-2"
[4] Limite de cortes (opcional): 5 [ou Enter para sem limite]
[5] Quantidade de bobinas: 4
[6] Peso total do lote (kg): 48000
```

### Passo 5: Visualizar Resultados

**Terminal:**
```
============================================================
  PLANO DE CORTE — COMBINAÇÕES VÁLIDAS
============================================================
  Âncora         : 50,80-2"
  Espessura      : 2.0 mm
  Tipo material  : COMERCIAL
  Largura bobina : 1200 mm
  Refilo mínimo  : 10 mm
  Combinações    : 104 (102 válidas + 2 fora da regra)

  #   Combinação                           Soma (mm)  Perda (mm)  Perda (%)  Status
  1   50,80-2"(x1) + 300X75X25(x2) + ...  1190.00    10.000      0.8333%    ✓ Válida
  ...
```

**Excel:**
- Arquivo salvo em: `output/plano_50-80-2in_esp2-0_COMERCIAL_L1200_20240315_143022.xlsx`

---

## Estrutura do Código

O código está organizado em **9 blocos funcionais**:

```python
# BLOCO 1: Configurações e Constantes
LARGURAS_BOBINA = [1200, 1000, 1500]
PERDA_MIN_PCT = 0.67
PERDA_MAX_PCT = 1.70
...

# BLOCO 2: Carga e Limpeza de Dados
def carregar_dados(caminho: str) -> pd.DataFrame

# BLOCO 3: Consultas ao Banco
def listar_espessuras(df: pd.DataFrame) -> list[float]
def listar_tipos(df: pd.DataFrame, espessura: float) -> list[str]
def listar_matrizes(df: pd.DataFrame, espessura: float, tipo: str) -> pd.DataFrame
def obter_desenvolvimento(df: pd.DataFrame, matriz: str, espessura: float) -> float

# BLOCO 4: Motor Combinatorial
def buscar_combinacoes_para_largura(...) -> list[dict]
def encontrar_combinacoes(...) -> tuple[pd.DataFrame, int]

# BLOCO 5: Cálculo de KG
def calcular_peso_medio_bobina(peso_informado: float, qtd_bobinas: int) -> float
def calcular_kg_matriz(...) -> float
def calcular_kg_combinacao(...) -> float

# BLOCO 6: Validação
def validar_resultado(df_res: pd.DataFrame, espessura: float) -> dict

# BLOCO 7: Interface CLI
def exibir_terminal(...) -> None
def menu_usuario(df: pd.DataFrame) -> tuple

# BLOCO 8: Exportação Excel
def exportar_excel(...) -> None

# BLOCO 9: Main
def main() -> None
```

Cada função possui:
- Documentação completa em português
- Type hints (tipagem)
- Descrição de entradas e saídas
- Comentários linha a linha nos algoritmos complexos

---

## Exemplos

### Exemplo 1: Espessura Fina (≤3mm) - Refilo 10mm

**Entrada:**
- Espessura: 2.0 mm
- Matriz âncora: `50,80-2"`
- Limite: 5 cortes

**Resultado:**
```
✓ 102 combinações válidas
⚠ 2 combinações fora da regra (perda < 10mm)

Melhor combinação:
50,80-2"(x1) + 300X75X25(x2) + 21,30-1/2"PT(x1)
Perda: 10.0 mm (0.83%) | KG: 1.891 kg | ✓ Válida
```

### Exemplo 2: Espessura Grossa (>3mm) - Refilo 14mm

**Entrada:**
- Espessura: 4.75 mm
- Matriz âncora: `152,40-6"`
- Sem limite de cortes

**Resultado:**
```
✓ 15 combinações válidas
⚠ 6 combinações fora da regra (perda < 14mm)

Exemplo de combinação fora da regra:
152,40-6"(x1) + 101,60-4"(x1) + 75X40(x3)
Perda: 9.0 mm (0.75%) | ⚠ Fora da regra
```

### Exemplo 3: Cálculo de KG

**Cenário:**
- 4 bobinas
- Peso total: 48.000 kg
- Peso médio: 12.000 kg/bobina
- Largura: 1.200 mm

**Combinação:**
- `50,80-2"(x3)`: dev=157mm → KG = (12.000/1.200) × (3 × 157 × 4) = **18.840 kg**

---

## ⚙️ Parâmetros de Configuração

Todos os parâmetros ficam no **BLOCO 1** do código:

```python
# Larguras de bobina (ordem de preferência)
LARGURAS_BOBINA = [1200, 1000, 1500]

# Janela de perda aceitável (%)
PERDA_MIN_PCT = 0.67
PERDA_MAX_PCT = 1.70

# Refilo mínimo por espessura (mm)
REFILO_MIN_ATE_3MM   = 10    # para esp ≤ 3.0 mm
REFILO_MIN_ACIMA_3MM = 14    # para esp > 3.0 mm

# Máximo de matrizes complementares
MAX_COMP_NA_COMBO = 2        # 3+ causa explosão combinatória

# Valores padrão para KG
PESO_MEDIO_BOB_PAD = 12_000  # kg (12 ton)
QTD_BOBINAS_PAD = 1
```

### Como Ajustar

| Para mudar | Edite | Exemplo |
|------------|-------|---------|
| Janela de perda | `PERDA_MIN_PCT` / `PERDA_MAX_PCT` | `0.50` / `2.00` |
| Refilo mínimo | `REFILO_MIN_ATE_3MM` / `REFILO_MIN_ACIMA_3MM` | `12` / `16` |
| Mais complementares | `MAX_COMP_NA_COMBO` | `3` (⚠️ lento) |
| Novas larguras | `LARGURAS_BOBINA` | `[1200, 1000, 1500, 1800]` |

---

## Regras de Negócio

### 1. Larguras de Bobina
- Fixas: **1000, 1200 ou 1500 mm**
- Tentadas em ordem: 1200 → 1000 → 1500
- Para na primeira com resultados válidos

### 2. Matriz Âncora
- **Sempre presente** (N ≥ 1 corte)
- Escolhida pelo usuário no menu
- Determina espessura e tipo da combinação

### 3. Matrizes Complementares
- Mesma espessura + mesmo tipo da âncora
- Até 2 complementares por combinação (padrão)
- Preenchem espaço restante da bobina

### 4. Validação em Cascata

**1º nível - Perda %:**
```
0,67% ≤ Perda% ≤ 1,70%
```

**2º nível - Refilo mínimo:**
```
Espessura ≤ 3.0 mm  →  Perda ≥ 10 mm
Espessura > 3.0 mm  →  Perda ≥ 14 mm
```

**3º nível - Limite de cortes (opcional):**
```
Σ N_i ≤ Limite_máquina
```

**Status final:**
- **Válida**: passou em todos os níveis
- **Fora da regra**: passou no % mas não no refilo
- **Rejeitada**: não passou no % ou limite (não aparece)

### 5. Ordenação dos Resultados

Combinações são ordenadas por:
1. Menor perda % (mais eficiente)
2. Menor N de cortes da âncora
3. Menor número de complementares

---
## Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

---

## Autor

**Jefersson Souza**

- GitHub: [@jeferssonsouza](https://github.com/jefersondev18)
- Empresa: Açotel Indústria e Comércio LTDA
