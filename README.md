# Controle de Investimento: simulador de fundos imobiliários (FIIs)

Ferramenta em Excel, com cara de aplicativo, que simula o investimento mensal em **fundos imobiliários**: do aporte de cada mês até o patrimônio acumulado e os dividendos que ele rende. A ferramenta também projeta o patrimônio em cenários de 2 a 40 anos e divide o aporte entre seis tipos de FII, conforme o perfil de investidor escolhido numa lista.

Projeto de prática do desafio de Excel da DIO, com o Expert Felipe Aguiar ("Felipão").

- **Arquivo:** `Controle_de_Investimento.xlsx`
- **Funções principais:** `VF` (em inglês, `FV`), `PROCV` (em inglês, `VLOOKUP`), validação de dados e intervalos nomeados
- **Dados:** não há base para baixar. Todos os valores e percentuais estão na própria planilha e são **de exemplo**.

> ⚠️ **Aviso:** o salário, as taxas e os percentuais por perfil são exemplos didáticos. Nada aqui é recomendação de investimento.

---

## 1. Como usar

1. Abra `Controle_de_Investimento.xlsx` no Excel e ative a edição.
2. Na aba visível, preencha as células **sem cor de fundo** (as cinzas são calculadas e ficam bloqueadas):
   - Em **CONFIGURAÇÕES**: salário e rendimento da carteira.
   - Em **INVESTIMENTO MENSAL**: quanto investir por mês, por quantos anos e a taxa de rendimento mensal.
3. Escolha o **perfil** (Conservador, Moderado ou Agressivo) na lista suspensa de **Perfil ➡️➡️➡️**.
4. Leia as respostas nas células cinzas e a divisão do aporte na tabela de tipos de fundo.

---

## 2. As cinco perguntas e onde estão as respostas

Aba `Planilha1`, bloco **INVESTIMENTO MENSAL**:

| # | Pergunta | Célula | Tipo | Valor no arquivo |
|---|----------|--------|------|------------------|
| 1 | Quanto investir por mês? | `D19` | Entrada | R$ 500,00 |
| 2 | Por quantos anos? | `D20` | Entrada | 5 |
| 3 | Qual a taxa de rendimento mensal? | `D21` | Entrada | 1,08% |
| 4 | Quanto de patrimônio vai acumular? | `D22` | **Calculado** | R$ 41.902,01 |
| 5 | Quanto vai receber de dividendos por mês? | `D23` | **Calculado** | R$ 372,93 |

---

## 3. Como o `VF` entra nos cálculos

### Patrimônio acumulado (`D22`)

```excel
=VF(taxa_mensal; periodo_investido*12; aporte*-1)
```

O `VF` (valor futuro) calcula quanto vale, no fim do prazo, uma série de aportes iguais feitos todo mês e corrigidos por uma taxa:

| Argumento | O que é | Na planilha |
|-----------|---------|-------------|
| `taxa` | Rendimento por período | `taxa_mensal` (1,08% ao mês) |
| `nper` | Número de períodos | `periodo_investido*12`, que converte anos em meses (5 × 12 = 60) |
| `pgto` | Valor aplicado a cada período | `aporte*-1`. O sinal negativo é necessário porque o Excel trata o aporte como dinheiro que **sai**; com o `-1` o resultado vem positivo |

### Dividendos mensais (`D23`)

```excel
=patrimonio*rendimento_carteira
```

Os dividendos são o patrimônio acumulado multiplicado pelo rendimento mensal da carteira (0,89% no exemplo): R$ 41.902,01 × 0,89% = **R$ 372,93 por mês**.

> A planilha tem **duas taxas diferentes**. `taxa_mensal` (`D21`) faz o patrimônio crescer no `VF`. `rendimento_carteira` (`D15`) define quanto desse patrimônio volta em dividendos. Elas são entradas independentes.

### Cenários de 2 a 40 anos (`A26:D33`)

A mesma conta é repetida para vários prazos. O prazo em anos fica na coluna `A`, e a fórmula usa referências absolutas (`$D$21`, `$D$19`) para não se perder ao arrastar:

```excel
C26: =VF($D$21; $A26*12; $D$19*-1)     ← patrimônio no prazo da linha
D26: =C26*rendimento_carteira          ← dividendos mensais nesse patrimônio
```

Resultados com aporte de R$ 500 e taxa de 1,08% ao mês:

| Prazo | Patrimônio | Dividendos mensais |
|------:|-----------:|-------------------:|
| 2 anos | R$ 13.615,43 | R$ 121,18 |
| 5 anos | R$ 41.902,01 | R$ 372,93 |
| 10 anos | R$ 121.728,83 | R$ 1.083,39 |
| 15 anos | R$ 273.805,59 | R$ 2.436,87 |
| 20 anos | R$ 563.524,50 | R$ 5.015,37 |
| 30 anos | R$ 2.166.952,41 | R$ 19.285,88 |
| 35 anos | R$ 4.170.127,09 | R$ 37.114,13 |
| 40 anos | R$ 7.986.341,93 | R$ 71.078,44 |

---

## 4. Como o `PROCV` entra nos cálculos

A divisão do aporte por tipo de fundo usa uma **chave composta**, formada por *perfil* + *tipo de fundo*.

**Tabela de apoio** (aba oculta `Planilha2`, `A3:D21`):

| Coluna | Conteúdo |
|--------|----------|
| `A` CHAVE | `=B4&"-"&C4`, por exemplo `Moderado-PAPEL` |
| `B` PERFIL | Conservador, Moderado ou Agressivo |
| `C` TIPO DE FUNDO | PAPEL, TIJOLO, HÍBRIDO, FUNDO DE FUNDOS, DESENVOLVIMENTO, HOTELARIA |
| `D` % | Percentual do aporte para aquele tipo |

**Na tela** (`C39:C44`), o percentual vem com:

```excel
=PROCV($D$35&"-"&B39; Planilha2!$A:$D; 4; FALSO)
```

1. `$D$35&"-"&B39` monta a chave a partir do perfil escolhido e do tipo da linha (por exemplo, `Moderado-PAPEL`).
2. O `PROCV` procura essa chave na coluna `A` da tabela de apoio.
3. O argumento `4` devolve a coluna `%` e o `FALSO` exige correspondência exata.
4. O valor em reais é `=C39*aporte` (coluna `D`), e o `TOTAL` (`D45`) soma os seis valores e deve igualar o aporte.

Trocar o perfil em `D35` muda a chave e, com ela, toda a divisão do aporte.

A validação de dados em `D35` é uma lista fixa: `Conservador, Moderado, Agressivo`.

---

## 5. Intervalos nomeados

| Nome | Refere-se a | Conteúdo |
|------|-------------|----------|
| `aporte` | `Planilha1!$D$19` | Quanto investir por mês |
| `periodo_investido` | `Planilha1!$D$20` | Prazo em anos |
| `taxa_mensal` | `Planilha1!$D$21` | Taxa de rendimento mensal usada no `VF` |
| `patrimonio` | `Planilha1!$D$22` | Patrimônio acumulado (resultado do `VF`) |
| `rendimento_carteira` | `Planilha1!$D$15` | Rendimento mensal da carteira, usado nos dividendos |

---

## 6. Percentuais de cada perfil

| Tipo de fundo | Conservador | Moderado | Agressivo |
|---------------|:-----------:|:--------:|:---------:|
| PAPEL | 30% | 25% | 60% |
| TIJOLO | 45% | 40% | 5% |
| HÍBRIDO | 10% | 5% | 15% |
| FUNDO DE FUNDOS | 5% | 5% | 5% |
| DESENVOLVIMENTO | 5% | 15% | 10% |
| HOTELARIA | 5% | 10% | 5% |
| **Total** | **100%** | **100%** | **100%** |

Cada perfil soma 100%.

**Origem dos percentuais:** <!-- PREENCHER: explique de onde vieram (exemplo do Expert, pesquisa própria, etc.) -->

Os percentuais são só exemplos didáticos, e não recomendação de investimento.

### Mesma simulação em dois perfis

Aporte de R$ 500 por mês:

| Tipo de fundo | Moderado | Agressivo |
|---------------|---------:|----------:|
| PAPEL | R$ 125,00 | R$ 300,00 |
| TIJOLO | R$ 200,00 | R$ 25,00 |
| HÍBRIDO | R$ 25,00 | R$ 75,00 |
| FUNDO DE FUNDOS | R$ 25,00 | R$ 25,00 |
| DESENVOLVIMENTO | R$ 75,00 | R$ 50,00 |
| HOTELARIA | R$ 50,00 | R$ 25,00 |
| **Total** | **R$ 500,00** | **R$ 500,00** |

---

## 7. Estrutura da planilha

| Aba | Visibilidade | Conteúdo |
|-----|--------------|----------|
| `Planilha1` | Visível | A "tela" da ferramenta: banner, configurações, investimento mensal, cenários e divisão por perfil |
| `Planilha2` | Oculta | Tabela de apoio com chave, perfil, tipo de fundo e percentual |

Escolhas de formatação para dar cara de aplicativo:

- Banner com a marca **Jansen Investimento** no topo.
- Linhas de grade e títulos de linha e coluna ocultos.
- Células calculadas em cinza e **bloqueadas**. As células de entrada ficam desbloqueadas, então, ao ativar a proteção da planilha, só elas podem ser editadas.
- Valores em reais e percentuais formatados.

Para ver a aba oculta: clique com o botão direito em uma guia, depois **Reexibir**.

---

## 8. O que mudei em relação à ferramenta do Expert

Diferenças em relação ao roteiro do desafio:

- **Banner próprio:** logo "Jansen Investimento" criado para o projeto.
- **Mais cenários:** o roteiro pede 2, 5, 10, 20 e 30 anos; esta versão também projeta 15, 35 e 40 anos.
- **Sugestão de investimento:** o campo de sugestão (`D16`) está preenchido com R$ 1.500,00 (30% do salário de exemplo, R$ 5.000,00).

---

## 9. Aprendizados

O desenvolvimento deste projeto foi de suma importância, pois, é uma oportunidade de aplicar todos os conhecimentos trabalhados nas aulas do Felipão na DIO, especialmente na aplicação de formulas e edição de nomes de células.
Além da prática com planilhas, o projeto permitiu exercitar a estruturação de informações, a organização de dados e a criação de uma ferramenta com finalidade definida.

---

## 10. Autor!

* [@mauroajj-cpu](https://github.com/mauroajj-cpu)
* [Linkedin](https://www.linkedin.com/in/mauro-jansen-204488104/)
* [Gmail](mailto:mauroajj@gmail.com?subject=Olá,%20Contrate-me%21%21)
* [Whatsapp](https://wa.me/5598985947388)

```
