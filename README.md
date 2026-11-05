# HLM — Planilha de gestão do novo empreendimento

Planilha de gestão administrativa, financeira e comercial da construtora, feita a partir do modelo **HLM ONE** e com a identidade visual HLM (preto #040203, laranja #D45508, creme #DBD4CA, cinza quente #AEA79F; tipografia Space Grotesk e Space Mono).
Ela corrige os erros do HLM ONE e foi ampliada.

## Arquivos (`planilhas/`)

| Arquivo | Para quê |
|---|---|
| `HLM_Gestao_Novo_Empreendimento.xlsm` | **Versão principal (Excel).** Tem o botão **BAIXAR ACERTO (PDF)** na aba ACERTO MENSAL. |
| `HLM_Gestao_Novo_Empreendimento.xlsx` | Mesma planilha sem macro, para Google Planilhas / Excel online. Para baixar o acerto: Arquivo › Fazer download › PDF › "Página atual". |
| `HLM_ONE_Gestao.xlsm` *(só local, fora do Git)* | HLM ONE preenchido com os dados reais dos acertos. Não fica no repositório porque contém dados financeiros reais. |

## Abas

1. **PAINEL GERAL**: 12 indicadores (aportes, custo, caixa, contas a pagar, orçamento, avanço físico, custo projetado, custo/m², VGV, vendas, lucro, recebíveis), alertas de gestão e 8 gráficos.
2. **PAINEL COMERCIAL**: VGV, ticket médio, preço/m², recebíveis, inadimplência, viabilidade, gráficos e espelho de vendas.
3. **ACERTO MENSAL**: escolha o mês e veja os custos por categoria (oficial/especial), o caixa do mês, o rateio por sócio com a posição acumulada (deve aportar / crédito) e o detalhamento das despesas e da folha. Tem o botão para baixar em PDF.
4. **FLUXO DE CAIXA**: 30 meses a partir da data de início, com realizado, previsão (a pagar / a receber) e curva S planejado × realizado.
5. **ORÇAMENTO**: por etapa: orçado, realizado, comprometido, % financeiro × % físico, custo projetado (EAC) e status.
6. **VIABILIDADE**: VGV, comissões, impostos, custo projetado, lucro, margem, ROI, ponto de equilíbrio, exposição de caixa, lucro por sócio e tabela de sensibilidade.
7. **CONFIG**: dados do empreendimento, parâmetros (taxa de adm, preço/m², comissão, impostos), sócios e % de participação, categorias, etapas, equipe e fornecedores.
8. **DESPESAS · FOLHA PAG · APORTES · VENDAS · RECEBIMENTOS**: lançamentos, com listas suspensas e uma coluna de verificação automática em cada linha.
9. **AUDITORIA**: 16 verificações de consistência e conciliação entre as abas.
10. **GUIA**: como usar, como os dados se conectam e a lista das correções feitas em relação ao HLM ONE.

## HLM ONE (dados reais)

A versão com os dados reais (`planilhas/HLM_ONE_Gestao.xlsm`), o CSV de lançamentos (`dados/`) e os scripts que leem os PDFs (`scripts/hlm_one/`) ficam **fora do Git** (veja o `.gitignore`).
Se esses arquivos existirem localmente, o `gerar_planilha.py` também gera a versão do HLM ONE.

## Regerar

```bash
python3 scripts/gerar_planilha.py          # gera os 3 arquivos em planilhas/
```

Para alterar a macro, edite `scripts/vba/ModAcerto.bas` e recompile o `vbaProject.bin`. Isso exige o LibreOffice:

```bash
soffice --headless --accept="socket,host=localhost,port=2002;urp;" --norestore &
python3 scripts/gerar_vba.py 15            # 15 = número de abas
```
