# 📊 Simulador de Investimentos em Fundos Imobiliários (FIIs)

Documentação do projeto desenvolvido para o desafio prático de Excel da **Digital Innovation One (DIO)**. O objetivo é automatizar o planejamento financeiro e a alocação de carteira em Fundos Imobiliários com base em simulações dinâmicas e perfil de investidor.

---

## 🎯 Objetivos do Projeto

* **Simulação Financeira:** Calcular o acúmulo de patrimônio a longo prazo e a projeção de dividendos mensais.
* **Alocação Inteligente:** Distribuir os aportes de forma dinâmica de acordo com o perfil de risco do investidor (Conservador, Moderado ou Agressivo).
* **Automação no Excel:** Aplicar funções financeiras (`VF`), de busca (`PROCV` / `VLOOKUP`) e de agregação (`SOMA`).

---

## 🛠️ Estrutura e Funcionalidades da Planilha

A solução foi estruturada em duas abas integradas:

### 1. Aba Principal (`Planilha1`)

* **Configurações Gerais:**
* **Sugestão de Aporte:** Define a taxa de investimento mensal com base na renda (ex: `R$ 2.600,00 * 30% = R$ 780,00`).
* **Cálculo de Patrimônio:** Utiliza a função financeiro-matemática `=VF(Taxa_Mensal; Anos * 12; -Aporte)` para projetar o valor acumulado final.
* **Dividendos Estimados:** Multiplica o patrimônio acumulado pela taxa média esperada de rendimento mensal (`Patrimônio * Rendimento_Carteira`).


* **Análise de Cenários Temporais:**
* Projeção automática de patrimônio e renda passiva para horizontes de **2, 5, 10, 15, 20 e 30 anos**.


* **Alocação por Perfil de Investidor:**
* Divisão dinâmica do aporte mensal entre as categorias de FIIs: **Papel, Tijolo, Híbridos, FOFs, Desenvolvimento e Hotelaria**.



### 2. Aba de Parâmetros (`Planilha2`)

* Matriz de distribuição percentual por perfil (Conservador, Moderado e Agressivo).
* Utilização de chave combinada (`Perfil-TipoFII`) para permitir a busca via `PROCV` (`VLOOKUP`).

---

## 📈 Tabela Resumo dos Cenários Simulados

Com um aporte mensal padrão de **R$ 780,00** a uma taxa de **1,079% ao mês**:

| Tempo de Investimento | Patrimônio Acumulado (R$) | Rendimento Mensal Estimado (1% a.m.) |
| --- | --- | --- |
| **2 Anos** (24 meses) | R$ 21.328,45 | R$ 213,28 |
| **5 Anos** (60 meses) | R$ 65.263,18 | R$ 652,63 |
| **10 Anos** (120 meses) | R$ 188.790,52 | R$ 1.887,91 |
| **15 Anos** (180 meses) | R$ 422.378,34 | R$ 4.223,78 |
| **20 Anos** (240 meses) | R$ 864.081,62 | R$ 8.640,82 |
| **30 Anos** (360 meses) | R$ 3.398.243,15 | R$ 33.982,43 |

---

## 🧩 Principais Fórmulas Utilizadas

**Valor Futuro (Patrimônio Acumulado):**

```excel
=VF(Taxa_Mensal; Anos * 12; -Aporte)

```

**Busca Dinâmica de Alocação por Perfil:**

```excel
=PROCV($C$32 & "-" & Tipo_FII; Planilha2!$A$2:$D$20; 4; FALSO)

```

**Rendimento dos Dividendos:**

```excel
=Patrimonio_Acumulado * Rendimento_Carteira

```

---

## 💡 Aprendizados e Conclusões

1. **Juros Compostos na Prática:** O tempo é o fator de maior impacto no acúmulo de patrimônio, conforme demonstrado no crescimento exponencial da tabela de cenários.
2. **Automação com Intervalos Nomeados:** O uso de nomes configurados (`Aporte`, `Taxa_Mensal`, `Rendimento_Carteira`) facilitou a leitura e a manutenção das fórmulas no Excel.
3. **Tomada de Decisão Baseada em Dados:** A planilha permite que o investidor visualize com clareza o efeito do aumento dos aportes e do tempo na construção de sua independência financeira.
