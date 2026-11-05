# -*- coding: utf-8 -*-
"""Gerador da planilha de gestao do novo empreendimento HLM.

Gera:
  planilhas/HLM_Gestao_Novo_Empreendimento.xlsm   (com botao/macro para baixar o acerto em PDF)
  planilhas/HLM_Gestao_Novo_Empreendimento.xlsx   (sem macro - Google Planilhas / Excel online)
  planilhas/HLM_ONE_Gestao.xlsm   (opcional, local: HLM ONE com dados reais — não versionado)

Uso:  python3 scripts/gerar_planilha.py
"""
import csv
import datetime as dt
import os

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

AQUI = os.path.dirname(os.path.abspath(__file__))
SAIDA = os.path.join(os.path.dirname(AQUI), "planilhas")
VBA_BIN = os.path.join(AQUI, "vba", "vbaProject.bin")
DADOS_HLM_ONE = os.path.join(os.path.dirname(AQUI), "dados", "hlm_one_lancamentos.csv")

# ----------------------------------------------------------------------------
# Dimensoes
# ----------------------------------------------------------------------------
N_MESES = 30            # colunas de meses no fluxo de caixa
D_INI, D_FIM = 6, 1005  # DESPESAS (1000 linhas)
F_INI, F_FIM = 6, 605   # FOLHA PAG (600 linhas)
A_INI, A_FIM = 6, 405   # APORTES (400 linhas)
V_INI, V_FIM = 6, 45    # VENDAS (40 unidades)
R_INI, R_FIM = 6, 605   # RECEBIMENTOS (600 linhas)
N_CAT, N_ETAPA, N_EQUIPE, N_FORN, N_SOCIOS = 20, 20, 40, 100, 6
DET_LINHAS = 80         # linhas de detalhamento de despesas no acerto
FOL_LINHAS = 40         # linhas de folha no acerto

ABAS = ["PAINEL GERAL", "PAINEL COMERCIAL", "ACERTO MENSAL", "FLUXO DE CAIXA", "ORÇAMENTO",
        "VIABILIDADE", "CONFIG", "DESPESAS", "FOLHA PAG", "APORTES", "VENDAS", "RECEBIMENTOS",
        "AUDITORIA", "GUIA", "CALC"]

FX = "'FLUXO DE CAIXA'"
OR = "'ORÇAMENTO'"
AC = "'ACERTO MENSAL'"

CATEGORIAS = [
    ("Materiais e Insumos", "Material", "SIM"),
    ("Material Hidráulico e Elétrico", "Material", "SIM"),
    ("Empreitas e Serviços Terceirizados", "Serviço", "SIM"),
    ("Equipamentos, Ferramentas e Locações", "Equipamento", "SIM"),
    ("Fretes e Transporte", "Logística", "SIM"),
    ("Retiradas e Entulho (caçambas)", "Logística", "SIM"),
    ("Diárias e Alimentação", "Mão de obra", "SIM"),
    ("Concessionárias (água, luz, internet)", "Canteiro", "SIM"),
    ("Projetos e Consultorias", "Técnico", "SIM"),
    ("Taxas, Licenças, IPTU e Cartório", "Legal", "SIM"),
    ("Escritório e Administrativo", "Administrativo", "SIM"),
    ("Seguros e Garantias", "Legal", "SIM"),
    ("Encargos Trabalhistas (INSS, FGTS)", "Mão de obra", "SIM"),
    ("EPI, Uniformes e Saúde Ocupacional", "Mão de obra", "SIM"),
    ("Comissões e Corretagem", "Comercial", "NÃO"),
    ("Marketing e Vendas", "Comercial", "NÃO"),
    ("Impostos sobre Vendas", "Comercial", "NÃO"),
    ("Outros Gastos", "Outros", "SIM"),
]
ETAPAS = [
    ("01 Projetos e Legalização", "SIM"),
    ("02 Serviços Preliminares e Canteiro", "SIM"),
    ("03 Fundação", "SIM"),
    ("04 Estrutura", "SIM"),
    ("05 Alvenaria e Vedações", "SIM"),
    ("06 Cobertura e Impermeabilização", "SIM"),
    ("07 Instalações Hidrossanitárias", "SIM"),
    ("08 Instalações Elétricas", "SIM"),
    ("09 Revestimentos e Pisos", "SIM"),
    ("10 Esquadrias e Vidros", "SIM"),
    ("11 Pintura", "SIM"),
    ("12 Louças, Metais e Acabamentos", "SIM"),
    ("13 Área Externa e Paisagismo", "SIM"),
    ("14 Limpeza e Entrega", "SIM"),
    ("15 Administração da Obra", "SIM"),
    ("16 Comercial e Vendas", "NÃO"),
]
SOCIOS = [("MAURUS PAI", 0.25), ("MAURUS FILHO", 0.25), ("ROBERTO", 0.25), ("ROBINSON", 0.25)]
FORMAS = "PIX,Transferência,Boleto,Cartão,Dinheiro,Cheque"
STATUS_UN = ["DISPONÍVEL", "RESERVADA", "VENDIDA", "PERMUTA"]
TIPO_PARC = "Sinal / Entrada,Parcela mensal,Intermediária,Chaves,Financiamento bancário,Outros"

# ----------------------------------------------------------------------------
# Paleta — guia de marca HLM (preto, laranja, creme e cinza quente)
# ----------------------------------------------------------------------------
INK = "#040203"        # preto HLM
INK2 = "#1C1A19"       # preto suave (barras de seção)
BRAND = "#D45508"      # laranja HLM
BRAND_D = "#A83F03"    # laranja escuro
BRAND_L = "#FBE6D6"    # laranja claro (tints)
CREAM = "#DBD4CA"      # creme HLM
CREAM_L = "#F4F1EC"    # creme claro
WARM = "#AEA79F"       # cinza quente
WARM_D = "#6B5E52"     # cinza quente escuro
WINE = "#951D25"       # vinho (alertas)
NAVY = INK
NAVY2 = INK2
TEAL = INK
ORANGE = WARM          # cor das abas de lançamento
GOLD = BRAND_D
INPUT = "#FDF3EA"      # célula de preenchimento
AUTO = CREAM_L         # célula calculada
LINE = "#D9D2C8"
GREEN_BG, GREEN_FG = "#E4EBDF", "#2F5D3A"
RED_BG, RED_FG = "#F6DEDB", WINE
AMB_BG, AMB_FG = BRAND_L, BRAND_D
SERIES = [BRAND, INK, WARM, BRAND_D, WINE, WARM_D, CREAM, "#E8894D"]
F_TIT = "Space Grotesk"   # tipografia de títulos (marca)
F_MONO = "Space Mono"     # rótulos / kicker (marca)
F_TXT = "Arial"           # dados

MOEDA = '"R$" #,##0.00;[Red]-"R$" #,##0.00;"–"'
MOEDA0 = '"R$" #,##0;[Red]-"R$" #,##0;"–"'
PCT = '0.0%;[Red]-0.0%;"–"'
DATA = "dd/mm/yyyy"
MES = "mmm/yy"


def rng(sheet, col, r1, r2):
    s = "'%s'" % sheet if (" " in sheet or not sheet.isascii()) else sheet
    return "%s!$%s$%d:$%s$%d" % (s, col, r1, col, r2)


class F:
    """Fabrica de formatos."""

    def __init__(self, wb):
        self.wb = wb
        self.cache = {}

    def __call__(self, **kw):
        base = {"font_name": F_TXT, "font_size": 10, "valign": "vcenter", "font_color": INK}
        base.update(kw)
        key = tuple(sorted(base.items()))
        if key not in self.cache:
            self.cache[key] = self.wb.add_format(base)
        return self.cache[key]


def build(path, exemplo=False, macro=False, dados=None):
    exemplo = exemplo or bool(dados)
    wb = xlsxwriter.Workbook(path, {"strings_to_numbers": False})
    wb.set_calc_mode("auto")
    wb.set_properties({"title": "HLM - Gestão do Novo Empreendimento", "author": "HLM",
                       "company": "HLM", "comments": "Gerado por scripts/gerar_planilha.py"})
    fm = F(wb)
    if macro:
        wb.add_vba_project(VBA_BIN)
        wb.set_vba_name("ThisWorkbook")

    ws = {}
    for i, nome in enumerate(ABAS):
        w = wb.add_worksheet(nome)
        if macro:
            w.set_vba_name("Sheet%d" % (i + 1))
        w.hide_gridlines(2)
        ws[nome] = w

    # ---------------- formatos comuns
    f_title = fm(bold=True, font_size=18, font_color="white", bg_color=INK, indent=1, font_name=F_TIT,
                 bottom=5, bottom_color=BRAND)
    f_sub = fm(font_color=WARM_D, font_size=9, indent=1, font_name=F_MONO)
    f_sec = fm(bold=True, font_size=11, font_color=INK, bg_color="#FFFFFF", font_name=F_TIT,
               bottom=2, bottom_color=BRAND, valign="bottom")
    f_hdr = fm(bold=True, font_color=INK, bg_color=CREAM, align="center", text_wrap=True, border=1, border_color="#C8BFB3",
               font_name=F_TIT, font_size=9)
    f_hdr_l = fm(bold=True, font_color=INK, bg_color=CREAM, align="left", text_wrap=True, border=1, border_color="#C8BFB3", indent=1,
                 font_name=F_TIT, font_size=9)
    f_hdr_mes = fm(bold=True, font_color=INK, bg_color=CREAM, align="center", border=1, border_color="#C8BFB3", num_format=MES,
                   font_name=F_TIT, font_size=9)

    def f_in(nf=None, **kw):
        d = dict(bg_color=INPUT, border=1, border_color=LINE)
        if nf:
            d["num_format"] = nf
        d.update(kw)
        return fm(**d)

    def f_au(nf=None, **kw):
        d = dict(bg_color=AUTO, border=1, border_color=LINE)
        if nf:
            d["num_format"] = nf
        d.update(kw)
        return fm(**d)

    f_lbl = fm(border=1, border_color=LINE, indent=1)
    f_lbl_b = fm(border=1, border_color=LINE, indent=1, bold=True)
    f_note = fm(italic=True, font_color=WARM_D, font_size=9, text_wrap=True, valign="top")
    f_tot_l = fm(bold=True, bg_color="#ECE7E0", border=1, border_color=LINE, indent=1)
    f_tot = fm(bold=True, bg_color="#ECE7E0", border=1, border_color=LINE, num_format=MOEDA)
    f_tot_p = fm(bold=True, bg_color="#ECE7E0", border=1, border_color=LINE, num_format=PCT)
    f_big_l = fm(bold=True, font_color="white", bg_color=INK, border=1, border_color=INK, indent=1, font_name=F_TIT)
    f_big = fm(bold=True, font_color=BRAND, bg_color=INK, border=1, border_color=INK, num_format=MOEDA, font_size=11)

    def nome_titulo(prefixo):
        return '=UPPER(NomeEmp)&"  |  %s"' % prefixo

    def cf_check(w, area):
        w.conditional_format(area, {"type": "cell", "criteria": "==", "value": '"OK"',
                                    "format": fm(font_color=GREEN_FG, bg_color=GREEN_BG, bold=True)})
        w.conditional_format(area, {"type": "formula", "criteria": "=AND(%s<>\"\",%s<>\"OK\")" % (area.split(":")[0].replace("$", ""), area.split(":")[0].replace("$", "")),
                                    "format": fm(font_color=RED_FG, bg_color=RED_BG, bold=True)})

    # =========================================================================
    # CONFIG
    # =========================================================================
    c = ws["CONFIG"]
    c.set_tab_color(ORANGE)
    c.set_column("A:A", 2)
    c.set_column("B:B", 44)
    c.set_column("C:C", 24)
    c.set_column("D:D", 46)
    c.set_column("E:E", 3)
    c.set_column("F:F", 38)
    c.set_column("G:G", 15)
    c.set_column("H:H", 11)
    c.set_column("I:I", 3)
    c.set_column("J:J", 36)
    c.set_column("K:K", 12)
    c.set_column("L:L", 3)
    c.set_column("M:M", 28)
    c.set_column("N:N", 20)
    c.set_column("O:O", 14)
    c.set_column("P:P", 3)
    c.set_column("Q:Q", 34)
    c.set_column("R:R", 28)
    c.set_row(1, 30)
    c.merge_range("B2:R2", "", f_title)
    c.write_formula("B2", nome_titulo("CONFIGURAÇÕES E CADASTROS"), f_title)
    c.merge_range("B3:R3", "Preencha as células LARANJA-CLARAS. Tudo o que está aqui alimenta automaticamente as demais abas (listas suspensas, rateio, taxas, viabilidade).", f_sub)

    c.merge_range("B4:D4", "DADOS DO EMPREENDIMENTO", f_sec)
    cfg = [
        # nome, label, valor(es), formato, nota, is_input
        ("NomeEmp", "Nome do empreendimento", "NOVO EMPREENDIMENTO HLM", None, "Ainda sem nome definitivo — altere aqui e todos os títulos mudam.", True),
        ("Construtora", "Construtora / administradora", "HLM Construções", None, "Responsável pela administração da obra.", True),
        ("Endereco", "Endereço / localização", "", None, "", True),
        ("Matricula", "Matrícula / alvará / ART", "", None, "", True),
        ("DataInicio", "Início do controle financeiro (mês 1)", dt.date(2026, 11, 1), DATA, "O FLUXO DE CAIXA começa neste mês e cobre %d meses." % N_MESES, True),
        ("Prazo", "Prazo previsto de obra (meses)", 18, "0", "Usado na curva de desembolso planejado e na data de entrega.", True),
        ("Entrega", "Previsão de entrega", "=EDATE(DataInicio,Prazo)", DATA, "Calculado.", False),
        ("AreaTerreno", "Área do terreno (m²)", 0, "#,##0.00", "", True),
        ("AreaConstruida", "Área construída total (m²)", 0, "#,##0.00", "Base para custo por m² construído.", True),
        ("AreaPrivativa", "Área privativa total (m²)", "=SUM(V_AREA)", "#,##0.00", "Calculado a partir da aba VENDAS.", False),
        ("NumUnid", "Número de unidades", "=COUNTIF(V_UNID,\"?*\")", "0", "Calculado a partir da aba VENDAS.", False),
        ("Conta", "Conta bancária oficial", "Sicredi — Conta Oficial", None, "", True),
    ]
    row = 4
    for nome, lbl, val, nf, nota, is_in in cfg:
        c.write(row, 1, lbl, f_lbl)
        fmt = f_in(nf) if is_in else f_au(nf)
        if isinstance(val, str) and val.startswith("="):
            c.write_formula(row, 2, val, fmt)
        elif isinstance(val, dt.date):
            c.write_datetime(row, 2, dt.datetime.combine(val, dt.time()), fmt)
        else:
            c.write(row, 2, val, fmt)
        c.write(row, 3, nota, f_note)
        wb.define_name(nome, "=CONFIG!$C$%d" % (row + 1))
        row += 1

    row += 1
    c.merge_range(row, 1, row, 3, "PARÂMETROS FINANCEIROS E COMERCIAIS", f_sec)
    row += 1
    params = [
        ("TaxaAdm", "Taxa de administração (sobre os gastos)", 0.15, PCT, "Aplicada mês a mês sobre despesas pagas cuja categoria tem 'INCIDE ADM = SIM' (+ folha, se marcado abaixo)."),
        ("FolhaIncideAdm", "Folha de pagamento incide taxa de adm.?", "SIM", None, "SIM / NÃO"),
        ("PrecoM2", "Preço de venda alvo (R$/m² privativo)", 0, MOEDA, "Gera o preço de tabela de cada unidade (área × preço × fator)."),
        ("ComissaoPadrao", "Comissão de corretagem padrão", 0.05, PCT, "Usada quando a unidade não tiver % própria."),
        ("ImpostoVendas", "Impostos sobre a receita (RET / presumido)", 0.04, PCT, "Ex.: RET do patrimônio de afetação = 4%."),
        ("CustoTerreno", "Custo do terreno / valor da permuta (R$)", 0, MOEDA, "Entra na viabilidade (não no fluxo da obra)."),
        ("Tolerancia", "Tolerância de desvio do orçamento", 0.05, PCT, "Acima disso a etapa fica em ATENÇÃO / ESTOURADA."),
    ]
    for nome, lbl, val, nf, nota in params:
        c.write(row, 1, lbl, f_lbl)
        c.write(row, 2, val, f_in(nf))
        c.write(row, 3, nota, f_note)
        wb.define_name(nome, "=CONFIG!$C$%d" % (row + 1))
        if nome == "FolhaIncideAdm":
            c.data_validation(row, 2, row, 2, {"validate": "list", "source": ["SIM", "NÃO"]})
        row += 1

    row += 1
    c.merge_range(row, 1, row, 3, "SÓCIOS E PARTICIPAÇÃO (base do rateio do acerto)", f_sec)
    row += 1
    c.write_row(row, 1, ["SÓCIO", "PARTICIPAÇÃO", "CONTATO (e-mail / telefone)"], f_hdr)
    row += 1
    soc_r1 = row + 1
    for i in range(N_SOCIOS):
        nm, pc = SOCIOS[i] if i < len(SOCIOS) else ("", None)
        c.write(row, 1, nm, f_in())
        if pc is None:
            c.write_blank(row, 2, None, f_in(PCT))
        else:
            c.write(row, 2, pc, f_in(PCT))
        c.write_blank(row, 3, None, f_in())
        row += 1
    soc_r2 = row
    wb.define_name("Socios", "=CONFIG!$B$%d:$B$%d" % (soc_r1, soc_r2))
    wb.define_name("Part", "=CONFIG!$C$%d:$C$%d" % (soc_r1, soc_r2))
    c.write(row, 1, "TOTAL (deve ser 100%)", f_tot_l)
    c.write_formula(row, 2, "=SUM(Part)", f_tot_p)
    c.write_formula(row, 3, '=IF(ROUND(SUM(Part),4)=1,"OK","ATENÇÃO: a soma das participações precisa ser 100%")', f_au())
    cf_check(c, "D%d:D%d" % (row + 1, row + 1))
    wb.define_name("PartTotal", "=CONFIG!$C$%d" % (row + 1))
    SOC_ROWS = list(range(soc_r1, soc_r2 + 1))

    # listas
    c.merge_range("F4:H4", "CATEGORIAS DE DESPESA", f_sec)
    c.write_row("F5", ["CATEGORIA", "GRUPO", "INCIDE ADM?"], f_hdr)
    for i in range(N_CAT):
        r = 5 + i
        cat = CATEGORIAS[i] if i < len(CATEGORIAS) else ("", "", "")
        c.write(r, 5, cat[0], f_in())
        c.write(r, 6, cat[1], f_in())
        c.write(r, 7, cat[2], f_in(align="center"))
    c.data_validation(5, 7, 5 + N_CAT - 1, 7, {"validate": "list", "source": ["SIM", "NÃO"]})
    c.write(5 + N_CAT, 5, "Grupo 'Comercial' = fora do custo da obra (comissão, marketing, impostos). Não incide adm.", f_note)
    wb.define_name("Categorias", "=CONFIG!$F$6:$F$%d" % (5 + N_CAT))
    wb.define_name("CatGrupo", "=CONFIG!$G$6:$G$%d" % (5 + N_CAT))
    wb.define_name("CatAdm", "=CONFIG!$H$6:$H$%d" % (5 + N_CAT))
    _cats = [x[0] for x in CATEGORIAS]
    CAT_COMISSAO = "CONFIG!$F$%d" % (6 + _cats.index("Comissões e Corretagem"))
    CAT_IMPOSTO = "CONFIG!$F$%d" % (6 + _cats.index("Impostos sobre Vendas"))

    c.merge_range("J4:K4", "ETAPAS DA OBRA (orçamento)", f_sec)
    c.write_row("J5", ["ETAPA", "CUSTO DE OBRA?"], f_hdr)
    for i in range(N_ETAPA):
        r = 5 + i
        et = ETAPAS[i] if i < len(ETAPAS) else ("", "")
        c.write(r, 9, et[0], f_in())
        c.write(r, 10, et[1], f_in(align="center"))
    c.data_validation(5, 10, 5 + N_ETAPA - 1, 10, {"validate": "list", "source": ["SIM", "NÃO"]})
    wb.define_name("Etapas", "=CONFIG!$J$6:$J$%d" % (5 + N_ETAPA))
    wb.define_name("EtapaObra", "=CONFIG!$K$6:$K$%d" % (5 + N_ETAPA))

    c.merge_range("M4:O4", "EQUIPE (folha de pagamento)", f_sec)
    c.write_row("M5", ["NOME", "FUNÇÃO", "DIÁRIA PADRÃO"], f_hdr)
    for i in range(N_EQUIPE):
        r = 5 + i
        c.write_blank(r, 12, None, f_in())
        c.write_blank(r, 13, None, f_in())
        c.write_blank(r, 14, None, f_in(MOEDA))
    wb.define_name("Equipe", "=CONFIG!$M$6:$M$%d" % (5 + N_EQUIPE))
    wb.define_name("EquipeFuncao", "=CONFIG!$N$6:$N$%d" % (5 + N_EQUIPE))
    wb.define_name("EquipeDiaria", "=CONFIG!$O$6:$O$%d" % (5 + N_EQUIPE))

    c.merge_range("Q4:R4", "FORNECEDORES (opcional)", f_sec)
    c.write_row("Q5", ["FORNECEDOR", "CNPJ / CONTATO"], f_hdr)
    for i in range(N_FORN):
        c.write_blank(5 + i, 16, None, f_in())
        c.write_blank(5 + i, 17, None, f_in())
    wb.define_name("Fornecedores", "=CONFIG!$Q$6:$Q$%d" % (5 + N_FORN))
    c.freeze_panes(4, 0)

    # =========================================================================
    # FLUXO DE CAIXA (estrutura definida antes porque meses sao referencia global)
    # =========================================================================
    fx = ws["FLUXO DE CAIXA"]
    C0 = 2  # coluna C = primeiro mes
    CL = xl_col_to_name(C0 + N_MESES - 1)
    CALC_CL = CL
    wb.define_name("Meses", "=%s!$C$5:$%s$5" % (FX, CL))
    wb.define_name("FluxoIni", "=%s!$C$5" % FX)
    wb.define_name("FluxoFim", "=%s!$%s$5" % (FX, CL))

    # nomes de colunas de lancamentos
    def nm(nome, sheet, col, r1, r2):
        wb.define_name(nome, "=" + rng(sheet, col, r1, r2))

    for n_, col in [("D_DATA", "A"), ("D_MES", "B"), ("D_TIPO", "C"), ("D_CAT", "D"), ("D_ETAPA", "E"),
                    ("D_FORN", "F"), ("D_DESC", "G"), ("D_NF", "H"), ("D_FORMA", "I"), ("D_VALOR", "J"),
                    ("D_STATUS", "K"), ("D_VENC", "L"), ("D_BASEADM", "M"), ("D_CHECK", "N"), ("D_ORD", "O")]:
        nm(n_, "DESPESAS", col, D_INI, D_FIM)
    for n_, col in [("F_DATA", "A"), ("F_MES", "B"), ("F_QUINZ", "C"), ("F_NOME", "D"), ("F_FUNCAO", "E"),
                    ("F_ETAPA", "F"), ("F_CAIXA", "G"), ("F_DIAS", "H"), ("F_DIARIA", "J"), ("F_EXTRA", "K"),
                    ("F_DESC", "L"), ("F_TOTAL", "M"), ("F_CHECK", "O"), ("F_ORD", "P")]:
        nm(n_, "FOLHA PAG", col, F_INI, F_FIM)
    for n_, col in [("A_DATA", "A"), ("A_MES", "B"), ("A_SOCIO", "C"), ("A_VALOR", "F"), ("A_CHECK", "G")]:
        nm(n_, "APORTES", col, A_INI, A_FIM)
    for n_, col in [("V_UNID", "A"), ("V_AREA", "C"), ("V_TAB", "E"), ("V_STATUS", "F"), ("V_CLIENTE", "G"),
                    ("V_DATA", "J"), ("V_VALOR", "K"), ("V_COMIS", "N"), ("V_REC", "O"), ("V_ARECEBER", "P"),
                    ("V_CHECK", "R")]:
        nm(n_, "VENDAS", col, V_INI, V_FIM)
    for n_, col in [("R_DATA", "A"), ("R_MES", "B"), ("R_UNID", "C"), ("R_VALOR", "F"), ("R_STATUS", "G"),
                    ("R_CHECK", "J")]:
        nm(n_, "RECEBIMENTOS", col, R_INI, R_FIM)

    # =========================================================================
    # DESPESAS
    # =========================================================================
    d = ws["DESPESAS"]
    d.set_tab_color(ORANGE)
    widths = [11, 9, 11, 30, 30, 26, 34, 12, 14, 15, 11, 12, 13, 22, 6]
    for i, wdt in enumerate(widths):
        d.set_column(i, i, wdt)
    d.set_column("M:M", 13, None, {"hidden": False})
    d.set_column("O:O", 6, None, {"hidden": True})
    d.set_row(1, 30)
    d.merge_range("A2:N2", "", f_title)
    d.write_formula("A2", nome_titulo("DESPESAS DA OBRA"), f_title)
    d.merge_range("A3:N3", "Lance TODOS os gastos. O MÊS é calculado pela DATA (não digite). STATUS em branco = PAGO. "
                  "Contas futuras: STATUS = A PAGAR + VENCIMENTO; quando pagar, mude para PAGO e ajuste a DATA.", f_sub)
    hdr = ["DATA", "MÊS (auto)", "TIPO", "CATEGORIA", "ETAPA", "FORNECEDOR", "DESCRIÇÃO", "Nº NF / RECIBO",
           "FORMA PGTO", "VALOR (R$)", "STATUS", "VENCIMENTO", "BASE ADM (auto)", "VERIFICAÇÃO (auto)", "ord", "OBSERVAÇÃO / ORIGEM"]
    d.set_row(4, 30)
    d.write_row(4, 0, hdr, f_hdr)
    d.write_formula(3, 8, '="Total lançado:"', fm(bold=True, align="right"))
    d.write_formula(3, 9, "=SUM(D_VALOR)", fm(bold=True, num_format=MOEDA))
    d.write_formula(3, 12, '=COUNTIFS(D_CHECK,"?*",D_CHECK,"<>OK")', fm(bold=True, font_color=RED_FG, num_format='0" pendência(s)";;"sem pendências"'))
    for r in range(D_INI, D_FIM + 1):
        i = r - 1
        R = str(r)
        d.write_blank(i, 0, None, f_in(DATA))
        d.write_formula(i, 1, '=IF(A%s="","",DATE(YEAR(A%s),MONTH(A%s),1))' % (R, R, R), f_au(MES, align="center"))
        for col in (2, 3, 4, 5, 6, 7, 8):
            d.write_blank(i, col, None, f_in())
        d.write_blank(i, 9, None, f_in(MOEDA))
        d.write_blank(i, 10, None, f_in(align="center"))
        d.write_blank(i, 11, None, f_in(DATA))
        d.write_formula(i, 12, '=IF(N(J{0})=0,0,IF(IFERROR(INDEX(CatAdm,MATCH(D{0},Categorias,0)),"SIM")="NÃO",0,J{0}))'.format(R),
                        f_au(MOEDA))
        chk = ('=IF(AND(A{0}="",J{0}=""),"",IF(J{0}="","SEM VALOR",IF(A{0}="","SEM DATA",'
               'IF(OR(B{0}<FluxoIni,B{0}>FluxoFim),"FORA DO PERÍODO DO FLUXO",'
               'IF(AND(C{0}<>"OFICIAL",C{0}<>"ESPECIAL"),"TIPO INVÁLIDO",'
               'IF(D{0}="","SEM CATEGORIA",IF(ISNA(MATCH(D{0},Categorias,0)),"CATEGORIA INVÁLIDA",'
               'IF(E{0}="","SEM ETAPA",IF(ISNA(MATCH(E{0},Etapas,0)),"ETAPA INVÁLIDA",'
               'IF(AND(K{0}<>"",K{0}<>"PAGO",K{0}<>"A PAGAR"),"STATUS INVÁLIDO",'
               'IF(AND(K{0}="A PAGAR",L{0}=""),"SEM VENCIMENTO",'
               'IF(AND(INDEX(CatGrupo,MATCH(D{0},Categorias,0))="Comercial",INDEX(EtapaObra,MATCH(E{0},Etapas,0))="SIM"),"CATEG. COMERCIAL EM ETAPA DE OBRA",'
               '"OK"))))))))))))').format(R)
        d.write_formula(i, 13, chk, f_au(align="center", font_size=9))
        d.write_formula(i, 14, '=IF(AND(B{0}=SelMes,N(J{0})<>0,K{0}<>"A PAGAR"),COUNTIFS(B${1}:B{0},SelMes,J${1}:J{0},"<>",K${1}:K{0},"<>A PAGAR"),"")'.format(R, D_INI),
                        fm(font_color=WARM, font_size=8))
    last = D_FIM - 1
    d.data_validation(D_INI - 1, 2, last, 2, {"validate": "list", "source": ["OFICIAL", "ESPECIAL"],
                                              "input_title": "Tipo", "input_message": "OFICIAL = com nota\nESPECIAL = sem nota"})
    d.data_validation(D_INI - 1, 3, last, 3, {"validate": "list", "source": "=Categorias"})
    d.data_validation(D_INI - 1, 4, last, 4, {"validate": "list", "source": "=Etapas"})
    d.data_validation(D_INI - 1, 5, last, 5, {"validate": "list", "source": "=Fornecedores", "show_error": False})
    d.data_validation(D_INI - 1, 8, last, 8, {"validate": "list", "source": FORMAS.split(",")})
    d.data_validation(D_INI - 1, 10, last, 10, {"validate": "list", "source": ["PAGO", "A PAGAR"]})
    d.data_validation(D_INI - 1, 0, last, 0, {"validate": "date", "criteria": "between",
                                              "minimum": dt.date(2015, 1, 1), "maximum": dt.date(2045, 12, 31),
                                              "error_message": "Digite uma data válida (dd/mm/aaaa)."})
    d.data_validation(D_INI - 1, 9, last, 9, {"validate": "decimal", "criteria": ">=", "value": 0,
                                              "error_message": "Valor deve ser positivo."})
    cf_check(d, "N%d:N%d" % (D_INI, D_FIM))
    d.conditional_format("K%d:K%d" % (D_INI, D_FIM), {"type": "cell", "criteria": "==", "value": '"A PAGAR"',
                                                       "format": fm(font_color=AMB_FG, bg_color=AMB_BG, bold=True)})
    d.conditional_format("L%d:L%d" % (D_INI, D_FIM), {"type": "formula", "criteria": '=AND(K%d="A PAGAR",L%d<>"",L%d<TODAY())' % (D_INI, D_INI, D_INI),
                                                       "format": fm(font_color=RED_FG, bg_color=RED_BG, bold=True)})
    d.freeze_panes(5, 1)
    d.set_column("P:P", 48)
    for r_ in range(D_INI, D_FIM + 1):
        d.write_blank(r_ - 1, 15, None, f_in(font_size=9))
    d.autofilter(4, 0, D_FIM - 1, 15)

    # =========================================================================
    # FOLHA PAG
    # =========================================================================
    fo = ws["FOLHA PAG"]
    fo.set_tab_color(ORANGE)
    for i, wdt in enumerate([11, 9, 17, 26, 18, 28, 11, 7, 12, 12, 12, 12, 14, 24, 22, 6]):
        fo.set_column(i, i, wdt)
    fo.set_column("P:P", 6, None, {"hidden": True})
    fo.set_row(1, 30)
    fo.merge_range("A2:O2", "", f_title)
    fo.write_formula("A2", nome_titulo("FOLHA DE PAGAMENTO"), f_title)
    fo.merge_range("A3:O3", "Uma linha por colaborador por quinzena. Cadastre a equipe em CONFIG (função e diária padrão são puxadas automaticamente). "
                   "CAIXA em branco = ESPECIAL.", f_sub)
    fo.set_row(4, 30)
    fo.write_row(4, 0, ["DATA PGTO", "MÊS (auto)", "QUINZENA", "COLABORADOR", "FUNÇÃO (auto)", "ETAPA", "CAIXA", "DIAS",
                        "DIÁRIA ESPECÍFICA", "DIÁRIA APLICADA", "EXTRAS (+)", "VALES / DESC. (−)", "TOTAL (R$)",
                        "OBSERVAÇÃO", "VERIFICAÇÃO (auto)", "ord"], f_hdr)
    fo.write_formula(3, 11, '="Total folha:"', fm(bold=True, align="right"))
    fo.write_formula(3, 12, "=SUM(F_TOTAL)", fm(bold=True, num_format=MOEDA))
    for r in range(F_INI, F_FIM + 1):
        i = r - 1
        R = str(r)
        fo.write_blank(i, 0, None, f_in(DATA))
        fo.write_formula(i, 1, '=IF(A{0}="","",DATE(YEAR(A{0}),MONTH(A{0}),1))'.format(R), f_au(MES, align="center"))
        fo.write_blank(i, 2, None, f_in())
        fo.write_blank(i, 3, None, f_in())
        fo.write_formula(i, 4, '=IF(D{0}="","",IFERROR(INDEX(EquipeFuncao,MATCH(D{0},Equipe,0))&"",""))'.format(R), f_au())
        fo.write_blank(i, 5, None, f_in())
        fo.write_blank(i, 6, None, f_in(align="center"))
        fo.write_blank(i, 7, None, f_in("0.0"))
        fo.write_blank(i, 8, None, f_in(MOEDA))
        fo.write_formula(i, 9, '=IF(I{0}<>"",I{0},IF(D{0}="",0,IFERROR(N(INDEX(EquipeDiaria,MATCH(D{0},Equipe,0))),0)))'.format(R), f_au(MOEDA))
        fo.write_blank(i, 10, None, f_in(MOEDA))
        fo.write_blank(i, 11, None, f_in(MOEDA))
        fo.write_formula(i, 12, '=IF(AND(H{0}="",K{0}="",L{0}=""),"",N(H{0})*J{0}+N(K{0})-N(L{0}))'.format(R), f_au(MOEDA, bold=True))
        fo.write_blank(i, 13, None, f_in())
        chk = ('=IF(AND(A{0}="",D{0}="",H{0}="",K{0}=""),"",IF(A{0}="","SEM DATA",'
               'IF(OR(B{0}<FluxoIni,B{0}>FluxoFim),"FORA DO PERÍODO DO FLUXO",IF(D{0}="","SEM COLABORADOR",'
               'IF(ISNA(MATCH(D{0},Equipe,0)),"NÃO CADASTRADO NA EQUIPE",IF(AND(N(H{0})>0,J{0}=0),"SEM DIÁRIA",'
               'IF(AND(F{0}<>"",ISNA(MATCH(F{0},Etapas,0))),"ETAPA INVÁLIDA",'
               'IF(AND(G{0}<>"",G{0}<>"OFICIAL",G{0}<>"ESPECIAL"),"CAIXA INVÁLIDO",'
               'IF(N(M{0})<0,"TOTAL NEGATIVO","OK")))))))))').format(R)
        fo.write_formula(i, 14, chk, f_au(align="center", font_size=9))
        fo.write_formula(i, 15, '=IF(AND(B{0}=SelMes,M{0}<>""),COUNTIFS(B${1}:B{0},SelMes,M${1}:M{0},"<>"),"")'.format(R, F_INI),
                         fm(font_color=WARM, font_size=8))
    last = F_FIM - 1
    fo.data_validation(F_INI - 1, 2, last, 2, {"validate": "list", "source": ["1ª Quinzena (01-15)", "2ª Quinzena (16-fim)", "Mensal", "Avulso"]})
    fo.data_validation(F_INI - 1, 3, last, 3, {"validate": "list", "source": "=Equipe", "show_error": False})
    fo.data_validation(F_INI - 1, 5, last, 5, {"validate": "list", "source": "=Etapas"})
    fo.data_validation(F_INI - 1, 6, last, 6, {"validate": "list", "source": ["OFICIAL", "ESPECIAL"]})
    fo.data_validation(F_INI - 1, 0, last, 0, {"validate": "date", "criteria": "between",
                                               "minimum": dt.date(2015, 1, 1), "maximum": dt.date(2045, 12, 31)})
    cf_check(fo, "O%d:O%d" % (F_INI, F_FIM))
    fo.freeze_panes(5, 1)
    fo.autofilter(4, 0, F_FIM - 1, 14)

    # =========================================================================
    # APORTES
    # =========================================================================
    ap = ws["APORTES"]
    ap.set_tab_color(ORANGE)
    for i, wdt in enumerate([11, 9, 22, 15, 36, 15, 22]):
        ap.set_column(i, i, wdt)
    ap.set_column("H:H", 3)
    ap.set_column("I:I", 22)
    ap.set_column("J:N", 15)
    ap.set_row(1, 30)
    ap.merge_range("A2:N2", "", f_title)
    ap.write_formula("A2", nome_titulo("APORTES DOS SÓCIOS"), f_title)
    ap.merge_range("A3:N3", "Registre cada depósito de sócio na conta da obra. O resumo ao lado compara o que cada um aportou com a sua cota (participação × custo líquido).", f_sub)
    ap.set_row(4, 30)
    ap.write_row(4, 0, ["DATA", "MÊS (auto)", "SÓCIO", "FORMA", "DESCRIÇÃO", "VALOR (R$)", "VERIFICAÇÃO (auto)"], f_hdr)
    for r in range(A_INI, A_FIM + 1):
        i = r - 1
        R = str(r)
        ap.write_blank(i, 0, None, f_in(DATA))
        ap.write_formula(i, 1, '=IF(A{0}="","",DATE(YEAR(A{0}),MONTH(A{0}),1))'.format(R), f_au(MES, align="center"))
        ap.write_blank(i, 2, None, f_in())
        ap.write_blank(i, 3, None, f_in())
        ap.write_blank(i, 4, None, f_in())
        ap.write_blank(i, 5, None, f_in(MOEDA))
        ap.write_formula(i, 6, ('=IF(AND(A{0}="",F{0}=""),"",IF(F{0}="","SEM VALOR",IF(A{0}="","SEM DATA",'
                                'IF(OR(B{0}<FluxoIni,B{0}>FluxoFim),"FORA DO PERÍODO DO FLUXO",'
                                'IF(ISNA(MATCH(C{0},Socios,0)),"SÓCIO INVÁLIDO","OK")))))').format(R), f_au(align="center", font_size=9))
    ap.data_validation(A_INI - 1, 2, A_FIM - 1, 2, {"validate": "list", "source": "=Socios"})
    ap.data_validation(A_INI - 1, 3, A_FIM - 1, 3, {"validate": "list", "source": FORMAS.split(",")})
    ap.data_validation(A_INI - 1, 0, A_FIM - 1, 0, {"validate": "date", "criteria": "between",
                                                    "minimum": dt.date(2015, 1, 1), "maximum": dt.date(2045, 12, 31)})
    cf_check(ap, "G%d:G%d" % (A_INI, A_FIM))
    # resumo por socio
    ap.set_row(4, 30)
    ap.write_row(4, 8, ["SÓCIO", "PARTICIPAÇÃO", "TOTAL APORTADO", "COTA LÍQUIDA DEVIDA*", "POSIÇÃO (+crédito / −deve)", "% DO TOTAL APORTADO"], f_hdr)
    for k, sr in enumerate(SOC_ROWS):
        i = 5 + k
        R = str(i + 1)
        ap.write_formula(i, 8, '=IF(CONFIG!$B${0}="","",CONFIG!$B${0})'.format(sr), f_au(bold=True))
        ap.write_formula(i, 9, '=IF(I{0}="","",CONFIG!$C${1})'.format(R, sr), f_au(PCT))
        ap.write_formula(i, 10, '=IF(I{0}="","",SUMIFS(A_VALOR,A_SOCIO,I{0}))'.format(R), f_au(MOEDA))
        ap.write_formula(i, 11, '=IF(I{0}="","",J{0}*({1}!$B$24-{1}!$B$14))'.format(R, FX), f_au(MOEDA))
        ap.write_formula(i, 12, '=IF(I{0}="","",K{0}-L{0})'.format(R), f_au(MOEDA, bold=True))
        ap.write_formula(i, 13, '=IF(I{0}="","",IFERROR(K{0}/SUM($K$6:$K${1}),0))'.format(R, 5 + N_SOCIOS), f_au(PCT))
    tr = 5 + N_SOCIOS
    ap.write(tr, 8, "TOTAL", f_tot_l)
    ap.write_formula(tr, 9, "=SUM(J6:J%d)" % tr, f_tot_p)
    for col in (10, 11, 12):
        cl = xl_col_to_name(col)
        ap.write_formula(tr, col, "=SUM(%s6:%s%d)" % (cl, cl, tr), f_tot)
    ap.write_formula(tr, 13, "=SUM(N6:N%d)" % tr, f_tot_p)
    ap.merge_range(tr + 1, 8, tr + 3, 13, "* Cota líquida = participação × (custo total c/ taxa de adm − recebimentos de vendas), considerando o período do fluxo de caixa. "
                   "Posição positiva = sócio aportou acima da sua parte (crédito); negativa = precisa aportar.", f_note)
    ap.conditional_format("M6:M%d" % tr, {"type": "cell", "criteria": "<", "value": -0.005, "format": fm(font_color=RED_FG, bold=True, bg_color=RED_BG)})
    ap.freeze_panes(5, 0)
    ap.autofilter(4, 0, A_FIM - 1, 6)

    # =========================================================================
    # VENDAS (estoque de unidades)
    # =========================================================================
    v = ws["VENDAS"]
    v.set_tab_color(ORANGE)
    for i, wdt in enumerate([11, 16, 10, 9, 15, 13, 26, 20, 20, 11, 15, 10, 10, 13, 15, 15, 10, 26]):
        v.set_column(i, i, wdt)
    v.set_row(1, 30)
    v.merge_range("A2:R2", "", f_title)
    v.write_formula("A2", nome_titulo("VENDAS E ESTOQUE DE UNIDADES"), f_title)
    v.merge_range("A3:R3", "Cadastre todas as unidades. Preço de tabela = área × preço/m² (CONFIG) × fator. RECEBIDO vem da aba RECEBIMENTOS.", f_sub)
    v.set_row(4, 30)
    v.write_row(4, 0, ["UNIDADE", "TIPOLOGIA", "ÁREA PRIV. (m²)", "FATOR PREÇO", "PREÇO TABELA (auto)", "STATUS", "CLIENTE",
                       "CPF / CONTATO", "CORRETOR / IMOBILIÁRIA", "DATA VENDA", "VALOR DE VENDA", "DESC. s/ TABELA",
                       "COMISSÃO %", "COMISSÃO R$ (auto)", "RECEBIDO (auto)", "SALDO A RECEBER", "% RECEBIDO", "VERIFICAÇÃO (auto)"], f_hdr)
    for r in range(V_INI, V_FIM + 1):
        i = r - 1
        R = str(r)
        v.write_blank(i, 0, None, f_in(bold=True))
        v.write_blank(i, 1, None, f_in())
        v.write_blank(i, 2, None, f_in("#,##0.00"))
        v.write_blank(i, 3, None, f_in("0.00"))
        v.write_formula(i, 4, '=IF(A{0}="","",N(C{0})*PrecoM2*IF(D{0}="",1,D{0}))'.format(R), f_au(MOEDA))
        v.write_blank(i, 5, None, f_in(align="center"))
        for col in (6, 7, 8):
            v.write_blank(i, col, None, f_in())
        v.write_blank(i, 9, None, f_in(DATA))
        v.write_blank(i, 10, None, f_in(MOEDA))
        v.write_formula(i, 11, '=IF(OR(K{0}="",N(E{0})=0),"",1-K{0}/E{0})'.format(R), f_au(PCT))
        v.write_blank(i, 12, None, f_in(PCT))
        v.write_formula(i, 13, '=IF(N(K{0})=0,"",K{0}*IF(M{0}="",ComissaoPadrao,M{0}))'.format(R), f_au(MOEDA))
        v.write_formula(i, 14, '=IF(A{0}="","",SUMIFS(R_VALOR,R_UNID,A{0},R_STATUS,"RECEBIDO"))'.format(R), f_au(MOEDA))
        v.write_formula(i, 15, '=IF(N(K{0})=0,"",K{0}-O{0})'.format(R), f_au(MOEDA))
        v.write_formula(i, 16, '=IF(N(K{0})=0,"",O{0}/K{0})'.format(R), f_au(PCT))
        chk = ('=IF(A{0}="",IF(OR(F{0}<>"",K{0}<>""),"SEM IDENTIFICAÇÃO",""),IF(COUNTIF(V_UNID,A{0})>1,"UNIDADE DUPLICADA",'
               'IF(F{0}="","SEM STATUS",IF(N(C{0})=0,"SEM ÁREA",'
               'IF(AND(F{0}="VENDIDA",OR(N(K{0})=0,J{0}="",G{0}="")),"VENDA INCOMPLETA (valor/data/cliente)",'
               'IF(AND(N(K{0})>0,F{0}<>"VENDIDA",F{0}<>"RESERVADA"),"VALOR EM UNIDADE NÃO VENDIDA",'
               'IF(N(O{0})>N(K{0})+0.005,IF(N(K{0})=0,"RECEBIMENTO SEM VENDA","RECEBIDO MAIOR QUE O CONTRATO"),"OK")))))))').format(R)
        v.write_formula(i, 17, chk, f_au(align="center", font_size=9))
    v.data_validation(V_INI - 1, 5, V_FIM - 1, 5, {"validate": "list", "source": STATUS_UN})
    cf_check(v, "R%d:R%d" % (V_INI, V_FIM))
    for st, bg, fg in [("VENDIDA", GREEN_BG, GREEN_FG), ("RESERVADA", AMB_BG, AMB_FG), ("PERMUTA", "#ECE7E0", "#3A3633")]:
        v.conditional_format("F%d:F%d" % (V_INI, V_FIM), {"type": "cell", "criteria": "==", "value": '"%s"' % st,
                                                         "format": fm(bg_color=bg, font_color=fg, bold=True)})
    v.conditional_format("Q%d:Q%d" % (V_INI, V_FIM), {"type": "data_bar", "bar_color": WARM_D, "bar_solid": True,
                                                       "min_type": "num", "min_value": 0, "max_type": "num", "max_value": 1})
    tr = V_FIM
    v.write(tr, 0, "TOTAL", f_tot_l)
    v.write_formula(tr, 2, "=SUM(V_AREA)", fm(bold=True, bg_color="#ECE7E0", border=1, border_color=LINE, num_format="#,##0.00"))
    for col in (4, 10, 13, 14, 15):
        cl = xl_col_to_name(col)
        v.write_formula(tr, col, "=SUM(%s%d:%s%d)" % (cl, V_INI, cl, V_FIM), f_tot)
    v.write_formula(tr, 5, '=COUNTIF(V_STATUS,"VENDIDA")&" vendidas"', f_tot_l)
    v.freeze_panes(5, 1)
    v.autofilter(4, 0, V_FIM - 1, 17)

    # =========================================================================
    # RECEBIMENTOS
    # =========================================================================
    rc = ws["RECEBIMENTOS"]
    rc.set_tab_color(ORANGE)
    for i, wdt in enumerate([11, 9, 11, 26, 22, 15, 13, 15, 30, 24]):
        rc.set_column(i, i, wdt)
    rc.set_row(1, 30)
    rc.merge_range("A2:J2", "", f_title)
    rc.write_formula("A2", nome_titulo("RECEBIMENTOS DE VENDAS"), f_title)
    rc.merge_range("A3:J3", "Lance o plano de pagamento de cada venda (A RECEBER com a data de vencimento). Ao receber, mude para RECEBIDO e ajuste a data.", f_sub)
    rc.set_row(4, 30)
    rc.write_row(4, 0, ["DATA (venc./receb.)", "MÊS (auto)", "UNIDADE", "CLIENTE (auto)", "TIPO DE PARCELA", "VALOR (R$)",
                        "STATUS", "FORMA", "OBSERVAÇÃO", "VERIFICAÇÃO (auto)"], f_hdr)
    for r in range(R_INI, R_FIM + 1):
        i = r - 1
        R = str(r)
        rc.write_blank(i, 0, None, f_in(DATA))
        rc.write_formula(i, 1, '=IF(A{0}="","",DATE(YEAR(A{0}),MONTH(A{0}),1))'.format(R), f_au(MES, align="center"))
        rc.write_blank(i, 2, None, f_in())
        rc.write_formula(i, 3, '=IF(C{0}="","",IFERROR(INDEX(V_CLIENTE,MATCH(C{0},V_UNID,0))&"",""))'.format(R), f_au())
        rc.write_blank(i, 4, None, f_in())
        rc.write_blank(i, 5, None, f_in(MOEDA))
        rc.write_blank(i, 6, None, f_in(align="center"))
        rc.write_blank(i, 7, None, f_in())
        rc.write_blank(i, 8, None, f_in())
        chk = ('=IF(AND(A{0}="",F{0}=""),"",IF(F{0}="","SEM VALOR",IF(A{0}="","SEM DATA",'
               'IF(OR(B{0}<FluxoIni,B{0}>FluxoFim),"FORA DO PERÍODO DO FLUXO",'
               'IF(ISNA(MATCH(C{0},V_UNID,0)),"UNIDADE INEXISTENTE",'
               'IF(AND(G{0}<>"RECEBIDO",G{0}<>"A RECEBER"),"STATUS INVÁLIDO",'
               'IF(AND(G{0}="A RECEBER",A{0}<TODAY()),"VENCIDO (inadimplência?)","OK")))))))').format(R)
        rc.write_formula(i, 9, chk, f_au(align="center", font_size=9))
    rc.data_validation(R_INI - 1, 2, R_FIM - 1, 2, {"validate": "list", "source": "=V_UNID"})
    rc.data_validation(R_INI - 1, 4, R_FIM - 1, 4, {"validate": "list", "source": TIPO_PARC.split(",")})
    rc.data_validation(R_INI - 1, 6, R_FIM - 1, 6, {"validate": "list", "source": ["RECEBIDO", "A RECEBER"]})
    rc.data_validation(R_INI - 1, 7, R_FIM - 1, 7, {"validate": "list", "source": FORMAS.split(",") + ["Financiamento"]})
    cf_check(rc, "J%d:J%d" % (R_INI, R_FIM))
    rc.conditional_format("G%d:G%d" % (R_INI, R_FIM), {"type": "cell", "criteria": "==", "value": '"RECEBIDO"',
                                                        "format": fm(bg_color=GREEN_BG, font_color=GREEN_FG, bold=True)})
    rc.conditional_format("G%d:G%d" % (R_INI, R_FIM), {"type": "cell", "criteria": "==", "value": '"A RECEBER"',
                                                        "format": fm(bg_color=AMB_BG, font_color=AMB_FG, bold=True)})
    rc.freeze_panes(5, 1)
    rc.autofilter(4, 0, R_FIM - 1, 9)

    # =========================================================================
    # FLUXO DE CAIXA
    # =========================================================================
    fx.set_tab_color(TEAL)
    fx.set_column("A:A", 46)
    fx.set_column("B:B", 16)
    fx.set_column(C0, C0 + N_MESES - 1, 13)
    fx.set_row(1, 30)
    fx.merge_range(1, 0, 1, C0 + N_MESES - 1, "", f_title)
    fx.write_formula("A2", nome_titulo("FLUXO DE CAIXA MENSAL"), f_title)
    fx.merge_range(2, 0, 2, C0 + N_MESES - 1, "100% automático: APORTES + RECEBIMENTOS − DESPESAS − FOLHA − TAXA DE ADM. "
                   "Linhas 'corrigir' mostram lançamentos sem tipo/sócio válidos (veja AUDITORIA). Linha 38 (planejado) é editável.", f_sub)
    fx.write(4, 0, "DESCRIÇÃO", f_hdr_l)
    fx.write(4, 1, "TOTAL", f_hdr)
    fx.write_formula(4, C0, "=DATE(YEAR(DataInicio),MONTH(DataInicio),1)", f_hdr_mes)
    for k in range(1, N_MESES):
        fx.write_formula(4, C0 + k, "=EDATE(%s5,1)" % xl_col_to_name(C0 + k - 1), f_hdr_mes)
    fx.freeze_panes(5, 2)

    f_fx = f_au(MOEDA)
    f_fx_b = fm(bold=True, bg_color="#ECE7E0", border=1, border_color=LINE, num_format=MOEDA)
    f_fx_warn = fm(bg_color="#FBEDEB", border=1, border_color=LINE, num_format=MOEDA, font_color=RED_FG, italic=True)
    f_fx_in = f_in(MOEDA)

    def fx_sec(r, text):
        fx.merge_range(r - 1, 0, r - 1, C0 + N_MESES - 1, text, f_sec)

    def fx_line(r, label, fn, fmt=f_fx, lfmt=f_lbl, total="sum"):
        """fn(colLetter) -> formula string (sem '=')."""
        if label.startswith("="):
            fx.write_formula(r - 1, 0, label, lfmt)
        else:
            fx.write(r - 1, 0, label, lfmt)
        for k in range(N_MESES):
            cl = xl_col_to_name(C0 + k)
            fx.write_formula(r - 1, C0 + k, "=" + fn(cl, k), fmt)
        if total == "sum":
            fx.write_formula(r - 1, 1, "=SUM(C%d:%s%d)" % (r, CL, r), f_fx_b)
        elif total == "last":
            fx.write_formula(r - 1, 1, "=%s%d" % (CL, r), f_fx_b)
        elif total:
            fx.write_formula(r - 1, 1, total, f_fx_b)

    fx_sec(6, "ENTRADAS")
    for k, sr in enumerate(SOC_ROWS):
        r = 7 + k
        fx_line(r, '="   Aporte — "&IF(CONFIG!$B$%d="","(sócio %d — não cadastrado)",CONFIG!$B$%d)' % (sr, k + 1, sr),
                lambda cl, kk, sr=sr: 'IF(CONFIG!$B$%d="",0,SUMIFS(A_VALOR,A_SOCIO,CONFIG!$B$%d,A_MES,%s$5))' % (sr, sr, cl))
    # r = 13
    fx_line(13, "   Aportes sem sócio válido (corrigir)", lambda cl, k: "SUMIFS(A_VALOR,A_MES,%s$5)-SUM(%s7:%s12)" % (cl, cl, cl), f_fx_warn)
    fx_line(14, "   Recebimentos de vendas", lambda cl, k: 'SUMIFS(R_VALOR,R_MES,%s$5,R_STATUS,"RECEBIDO")' % cl)
    fx_line(15, "TOTAL DE ENTRADAS", lambda cl, k: "SUM(%s7:%s14)" % (cl, cl), f_fx_b, f_tot_l)
    fx_sec(16, "SAÍDAS (somente lançamentos PAGOS)")
    fx_line(17, "   Despesas — caixa OFICIAL (com nota)", lambda cl, k: 'SUMIFS(D_VALOR,D_MES,%s$5,D_TIPO,"OFICIAL",D_STATUS,"<>A PAGAR")' % cl)
    fx_line(18, "   Despesas — caixa ESPECIAL (sem nota)", lambda cl, k: 'SUMIFS(D_VALOR,D_MES,%s$5,D_TIPO,"ESPECIAL",D_STATUS,"<>A PAGAR")' % cl)
    fx_line(19, "   Despesas sem tipo válido (corrigir)", lambda cl, k: 'SUMIFS(D_VALOR,D_MES,%s$5,D_STATUS,"<>A PAGAR")-%s17-%s18' % (cl, cl, cl), f_fx_warn)
    fx_line(20, "   Folha de pagamento — caixa OFICIAL", lambda cl, k: 'SUMIFS(F_TOTAL,F_MES,%s$5,F_CAIXA,"OFICIAL")' % cl)
    fx_line(21, "   Folha de pagamento — caixa ESPECIAL", lambda cl, k: "SUMIFS(F_TOTAL,F_MES,%s$5)-%s20" % (cl, cl))
    fx_line(22, "CUSTO DIRETO DO MÊS", lambda cl, k: "SUM(%s17:%s21)" % (cl, cl), f_fx_b, f_tot_l)
    fx_line(23, '="   Taxa de administração ("&ROUND(TaxaAdm*100,2)&"%)"',
            lambda cl, k: 'TaxaAdm*(SUMIFS(D_BASEADM,D_MES,%s$5,D_STATUS,"<>A PAGAR")+IF(FolhaIncideAdm="SIM",%s20+%s21,0))' % (cl, cl, cl))
    fx_line(24, "TOTAL DE SAÍDAS", lambda cl, k: "%s22+%s23" % (cl, cl), f_fx_b, f_tot_l)
    fx_line(26, "SALDO DO MÊS", lambda cl, k: "%s15-%s24" % (cl, cl), fm(bold=True, border=1, border_color=LINE, num_format=MOEDA), f_lbl_b)
    fx_line(27, "SALDO ACUMULADO EM CAIXA",
            lambda cl, k: ("%s26" % cl) if k == 0 else ("%s27+%s26" % (xl_col_to_name(C0 + k - 1), cl)),
            f_big, f_big_l, total="last")
    fx_sec(28, "PREVISÃO (compromissos futuros pelo vencimento)")
    fx_line(29, "   Contas a pagar (c/ taxa de adm)",
            lambda cl, k: 'SUMIFS(D_VALOR,D_STATUS,"A PAGAR",D_VENC,">="&%s$5,D_VENC,"<"&EDATE(%s$5,1))+TaxaAdm*SUMIFS(D_BASEADM,D_STATUS,"A PAGAR",D_VENC,">="&%s$5,D_VENC,"<"&EDATE(%s$5,1))' % (cl, cl, cl, cl))
    fx_line(30, "   Recebimentos previstos (a receber)",
            lambda cl, k: 'SUMIFS(R_VALOR,R_STATUS,"A RECEBER",R_DATA,">="&%s$5,R_DATA,"<"&EDATE(%s$5,1))' % (cl, cl))
    fx_line(31, "SALDO PROJETADO ACUMULADO (realizado + previsto)",
            lambda cl, k: ("%s26+%s30-%s29" % (cl, cl, cl)) if k == 0 else ("%s31+%s26+%s30-%s29" % (xl_col_to_name(C0 + k - 1), cl, cl, cl)),
            fm(bold=True, border=1, border_color=LINE, num_format=MOEDA, bg_color=BRAND_L), f_lbl_b, total="last")
    fx_sec(33, "INDICADORES E CURVA DE DESEMBOLSO")
    fx_line(34, "   Custo realizado acumulado (c/ adm)",
            lambda cl, k: ("%s24" % cl) if k == 0 else ("%s34+%s24" % (xl_col_to_name(C0 + k - 1), cl)), total="last")
    fx_line(35, "   Aportes acumulados",
            lambda cl, k: ("SUM(%s7:%s13)" % (cl, cl)) if k == 0 else ("%s35+SUM(%s7:%s13)" % (xl_col_to_name(C0 + k - 1), cl, cl)), total="last")
    fx_line(36, "   Recebimentos de vendas acumulados",
            lambda cl, k: ("%s14" % cl) if k == 0 else ("%s36+%s14" % (xl_col_to_name(C0 + k - 1), cl)), total="last")
    fx_line(37, "   % do custo direto com nota (oficial)", lambda cl, k: "IFERROR((%s17+%s20)/%s22,0)" % (cl, cl, cl),
            f_au(PCT), total="=IFERROR((B17+B20)/B22,0)")
    # planejado (input com formula padrao)
    fx.write(37, 0, "   Desembolso PLANEJADO no mês (editável)", fm(border=1, border_color=LINE, indent=1, bold=True, font_color=GOLD))
    for k in range(N_MESES):
        fx.write_formula(37, C0 + k, "=IF(%d<=Prazo,%s!$B$29/Prazo,0)" % (k + 1, OR), f_fx_in)
    fx.write_formula(37, 1, "=SUM(C38:%s38)" % CL, f_fx_b)
    fx_line(39, "   Planejado acumulado",
            lambda cl, k: ("%s38" % cl) if k == 0 else ("%s39+%s38" % (xl_col_to_name(C0 + k - 1), cl)), total="last")
    fx_line(40, "   Aderência: realizado ÷ planejado (acumulado)", lambda cl, k: "IFERROR(%s34/%s39,0)" % (cl, cl), f_au(PCT), total="=IFERROR(B34/B39,0)")
    fx.write(41, 0, "A linha 38 vem preenchida com o orçamento projetado ÷ prazo (distribuição linear). Substitua pelos valores do seu cronograma físico-financeiro.", f_note)
    for rr in (26, 27, 31):
        fx.conditional_format(rr - 1, C0, rr - 1, C0 + N_MESES - 1, {"type": "cell", "criteria": "<", "value": -0.005,
                                                                      "format": fm(font_color="#FFFFFF", bg_color=WINE, bold=True)})
    for rr in (13, 19):
        fx.conditional_format(rr - 1, 1, rr - 1, C0 + N_MESES - 1, {"type": "cell", "criteria": "not between", "minimum": -0.005, "maximum": 0.005,
                                                                     "format": fm(bg_color="#E8B4AE", bold=True)})
    fx.set_landscape()
    fx.fit_to_pages(0, 1)
    fx.repeat_columns(0, 1)

    # =========================================================================
    # ORCAMENTO
    # =========================================================================
    o = ws["ORÇAMENTO"]
    o.set_tab_color(TEAL)
    for i, wdt in enumerate([36, 15, 8, 15, 14, 15, 14, 15, 11, 11, 11, 16, 15, 14, 9]):
        o.set_column(i, i, wdt)
    o.set_row(1, 30)
    o.merge_range("A2:O2", "", f_title)
    o.write_formula("A2", nome_titulo("ORÇAMENTO × REALIZADO × AVANÇO FÍSICO"), f_title)
    o.merge_range("A3:O3", "Preencha ORÇADO e % AVANÇO FÍSICO (células laranja-claras). O realizado vem de DESPESAS (pagas) e FOLHA PAG pela ETAPA. "
                   "Custo projetado (EAC) = realizado ÷ avanço físico (a partir de 10% de avanço).", f_sub)
    o.set_row(4, 36)
    o.write_row(4, 0, ["ETAPA", "ORÇADO (R$)", "PESO", "REALIZADO DESPESAS", "REALIZADO FOLHA", "REALIZADO TOTAL",
                       "COMPROMETIDO A PAGAR", "SALDO DO ORÇAMENTO", "% FINANCEIRO", "% AVANÇO FÍSICO", "DESVIO (fin − fís)",
                       "CUSTO PROJETADO (EAC)", "VARIAÇÃO PROJ. × ORÇ.", "STATUS", "CUSTO DE OBRA?"], f_hdr)
    O1, O2 = 6, 5 + N_ETAPA  # 6..25
    ONC = O2 + 1  # 26 nao classificado
    OSUB, OADM, OTOT = ONC + 1, ONC + 2, ONC + 3  # 27, 28, 29
    for r in range(O1, ONC + 1):
        i = r - 1
        R = str(r)
        if r == ONC:
            o.write(i, 0, "Sem etapa / etapa inválida (corrigir)", fm(border=1, border_color=LINE, indent=1, italic=True, font_color=RED_FG))
            o.write_blank(i, 1, None, f_au(MOEDA))
            o.write_formula(i, 3, '=SUMIFS(D_VALOR,D_STATUS,"<>A PAGAR")-SUM(D%d:D%d)' % (O1, O2), f_au(MOEDA))
            o.write_formula(i, 4, "=SUM(F_TOTAL)-SUM(E%d:E%d)" % (O1, O2), f_au(MOEDA))
            o.write_formula(i, 6, '=SUMIFS(D_VALOR,D_STATUS,"A PAGAR")-SUM(G%d:G%d)' % (O1, O2), f_au(MOEDA))
            o.write_blank(i, 9, None, f_au(PCT))
            o.write(i, 14, "SIM", f_au(align="center"))
        else:
            o.write_formula(i, 0, '=IF(CONFIG!$J${0}="","",CONFIG!$J${0})'.format(r), f_au(bold=True))
            o.write_blank(i, 1, None, f_in(MOEDA))
            o.write_formula(i, 3, '=IF(A{0}="",0,SUMIFS(D_VALOR,D_ETAPA,A{0},D_STATUS,"<>A PAGAR"))'.format(R), f_au(MOEDA))
            o.write_formula(i, 4, '=IF(A{0}="",0,SUMIFS(F_TOTAL,F_ETAPA,A{0}))'.format(R), f_au(MOEDA))
            o.write_formula(i, 6, '=IF(A{0}="",0,SUMIFS(D_VALOR,D_ETAPA,A{0},D_STATUS,"A PAGAR"))'.format(R), f_au(MOEDA))
            o.write_blank(i, 9, None, f_in(PCT))
            o.write_formula(i, 14, '=IF(A{0}="","",IF(CONFIG!$K${1}="NÃO","NÃO","SIM"))'.format(R, r), f_au(align="center"))
        o.write_formula(i, 2, "=IFERROR(B{0}/$B${1},0)".format(R, OSUB), f_au(PCT))
        o.write_formula(i, 5, "=D{0}+E{0}".format(R), f_au(MOEDA, bold=True))
        o.write_formula(i, 7, "=B{0}-F{0}-G{0}".format(R), f_au(MOEDA))
        o.write_formula(i, 8, "=IF(N(B{0})=0,IF(F{0}+G{0}>0,1,0),(F{0}+G{0})/B{0})".format(R), f_au(PCT))
        o.write_formula(i, 10, "=I{0}-N(J{0})".format(R), f_au(PCT))
        o.write_formula(i, 11, "=IF(N(J{0})>=0.1,MAX(F{0}+G{0},F{0}/J{0}),MAX(N(B{0}),F{0}+G{0}))".format(R), f_au(MOEDA))
        o.write_formula(i, 12, "=L{0}-B{0}".format(R), f_au(MOEDA))
        o.write_formula(i, 13, ('=IF(AND(N(B{0})=0,F{0}+G{0}=0),"—",IF(N(B{0})=0,"SEM ORÇAMENTO",'
                                'IF(I{0}>1+Tolerancia,"ESTOURADO",IF(OR(K{0}>Tolerancia,L{0}>B{0}*(1+Tolerancia)),"ATENÇÃO","OK"))))').format(R),
                        f_au(align="center", bold=True))
    R = str(OSUB)
    o.write(OSUB - 1, 0, "SUBTOTAL — CUSTO DIRETO", f_tot_l)
    for col in (1, 3, 4, 5, 6, 7, 11, 12):
        cl = xl_col_to_name(col)
        o.write_formula(OSUB - 1, col, "=SUM(%s%d:%s%d)" % (cl, O1, cl, ONC), f_tot)
    o.write_formula(OSUB - 1, 2, "=IFERROR(B%s/B%s,0)" % (R, R), f_tot_p)
    o.write_formula(OSUB - 1, 8, "=IFERROR((F{0}+G{0})/B{0},0)".format(R), f_tot_p)
    o.write_formula(OSUB - 1, 9, "=IFERROR(SUMPRODUCT(B%d:B%d,J%d:J%d)/B%s,0)" % (O1, O2, O1, O2, R), f_tot_p)
    o.write_formula(OSUB - 1, 10, "=I{0}-J{0}".format(R), f_tot_p)
    o.write_blank(OSUB - 1, 13, None, f_tot)
    o.write_blank(OSUB - 1, 14, None, f_tot)
    R = str(OADM)
    o.write_formula(OADM - 1, 0, '="Taxa de administração ("&ROUND(TaxaAdm*100,2)&"%)"', f_lbl_b)
    o.write_formula(OADM - 1, 1, '=TaxaAdm*SUMIFS(B{0}:B{1},O{0}:O{1},"SIM")'.format(O1, ONC), f_au(MOEDA))
    o.write_formula(OADM - 1, 5, "=%s!$B$23" % FX, f_au(MOEDA))
    o.write_formula(OADM - 1, 6, '=TaxaAdm*SUMIFS(D_BASEADM,D_STATUS,"A PAGAR")', f_au(MOEDA))
    o.write_formula(OADM - 1, 7, "=B{0}-F{0}-G{0}".format(R), f_au(MOEDA))
    o.write_formula(OADM - 1, 11, '=MAX(F{0}+G{0},TaxaAdm*SUMIFS(L{1}:L{2},O{1}:O{2},"SIM"))'.format(R, O1, ONC), f_au(MOEDA))
    o.write_formula(OADM - 1, 12, "=L{0}-B{0}".format(R), f_au(MOEDA))
    R = str(OTOT)
    o.write(OTOT - 1, 0, "TOTAL GERAL (c/ taxa de adm)", f_big_l)
    for col in (1, 5, 6, 7, 11, 12):
        cl = xl_col_to_name(col)
        o.write_formula(OTOT - 1, col, "=%s%d+%s%d" % (cl, OSUB, cl, OADM), f_big)
    o.write_formula(OTOT - 1, 8, "=IFERROR((F{0}+G{0})/B{0},0)".format(R), fm(bold=True, font_color="white", bg_color=TEAL, border=1, border_color=LINE, num_format=PCT))
    o.merge_range(OTOT, 0, OTOT + 1, 14, "Etapas com 'CUSTO DE OBRA? = NÃO' (ex.: Comercial e Vendas) entram no caixa e no acerto, mas ficam fora do custo da obra na VIABILIDADE "
                  "(comissões e impostos são calculados lá a partir das VENDAS). Peso e avanço físico ponderado usam o ORÇADO de cada etapa.", f_note)
    o.conditional_format("N%d:N%d" % (O1, ONC), {"type": "cell", "criteria": "==", "value": '"OK"', "format": fm(bg_color=GREEN_BG, font_color=GREEN_FG, bold=True)})
    o.conditional_format("N%d:N%d" % (O1, ONC), {"type": "cell", "criteria": "==", "value": '"ATENÇÃO"', "format": fm(bg_color=AMB_BG, font_color=AMB_FG, bold=True)})
    o.conditional_format("N%d:N%d" % (O1, ONC), {"type": "cell", "criteria": "==", "value": '"ESTOURADO"', "format": fm(bg_color=RED_BG, font_color=RED_FG, bold=True)})
    o.conditional_format("N%d:N%d" % (O1, ONC), {"type": "cell", "criteria": "==", "value": '"SEM ORÇAMENTO"', "format": fm(bg_color=RED_BG, font_color=RED_FG, bold=True)})
    o.conditional_format("I%d:J%d" % (O1, O2), {"type": "data_bar", "bar_color": BRAND, "bar_solid": True,
                                                 "min_type": "num", "min_value": 0, "max_type": "num", "max_value": 1})
    o.conditional_format("D%d:G%d" % (ONC, ONC), {"type": "cell", "criteria": "not between", "minimum": -0.005, "maximum": 0.005,
                                                   "format": fm(bg_color="#E8B4AE", bold=True)})
    o.data_validation(O1 - 1, 9, O2 - 1, 9, {"validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 1,
                                             "error_message": "Use um percentual entre 0% e 100%."})
    o.freeze_panes(5, 1)
    o.set_landscape()
    o.fit_to_pages(1, 0)

    # =========================================================================
    # VIABILIDADE
    # =========================================================================
    vi = ws["VIABILIDADE"]
    vi.set_tab_color(TEAL)
    vi.set_column("A:A", 2)
    vi.set_column("B:B", 50)
    vi.set_column("C:C", 18)
    vi.set_column("D:D", 46)
    vi.set_column("E:E", 3)
    vi.set_column("F:F", 26)
    vi.set_column("G:M", 14)
    vi.set_row(1, 30)
    vi.merge_range("B2:M2", "", f_title)
    vi.write_formula("B2", nome_titulo("VIABILIDADE E RESULTADO PROJETADO"), f_title)
    vi.merge_range("B3:M3", "Tudo calculado: VENDAS + ORÇAMENTO (custo projetado) + FLUXO + parâmetros de CONFIG.", f_sub)
    VI = {}

    def vrow(r, key, label, formula, nf=MOEDA, note="", strong=False):
        vi.write(r - 1, 1, label, f_tot_l if strong else f_lbl)
        vi.write_formula(r - 1, 2, formula, (f_tot if nf == MOEDA else fm(bold=True, bg_color="#ECE7E0", border=1, border_color=LINE, num_format=nf)) if strong else f_au(nf))
        if note:
            vi.write(r - 1, 3, note, f_note)
        if key:
            VI[key] = "VIABILIDADE!$C$%d" % r
            wb.define_name("vb_" + key, "=VIABILIDADE!$C$%d" % r)

    vi.merge_range("B5:D5", "RECEITA", f_sec)
    vrow(6, "vgv_tab", "VGV de tabela (todas as unidades, exceto permuta)", '=SUMIFS(V_TAB,V_STATUS,"<>PERMUTA")')
    vrow(7, "vgv_vend", "VGV contratado (unidades VENDIDAS)", '=SUMIFS(V_VALOR,V_STATUS,"VENDIDA")')
    vrow(8, "vgv_est", "VGV em estoque (disponíveis + reservadas, a preço de tabela)", '=SUMIFS(V_TAB,V_STATUS,"DISPONÍVEL")+SUMIFS(V_TAB,V_STATUS,"RESERVADA")')
    vrow(9, "vgv", "VGV TOTAL ESPERADO", "=C7+C8", strong=True)
    vrow(10, "comis", "(−) Comissões de corretagem", '=SUMIFS(V_COMIS,V_STATUS,"VENDIDA")+C8*ComissaoPadrao', note="Vendidas: % de cada unidade. Estoque: % padrão.")
    vrow(11, "imp", "(−) Impostos sobre a receita", "=C9*ImpostoVendas")
    vrow(12, "rec_liq", "RECEITA LÍQUIDA", "=C9-C10-C11", strong=True)
    vi.merge_range("B14:D14", "CUSTOS", f_sec)
    vrow(15, "orc", "Custo da obra ORÇADO (c/ adm)", '=SUMIFS(%s!$B$%d:$B$%d,%s!$O$%d:$O$%d,"SIM")+%s!$B$%d' % (OR, O1, ONC, OR, O1, ONC, OR, OADM))
    vrow(16, "real", "Custo realizado até hoje (c/ adm) — todo o caixa", "=%s!$B$24" % FX)
    vrow(17, "eac", "Custo da obra PROJETADO (EAC, c/ adm)", '=SUMIFS(%s!$L$%d:$L$%d,%s!$O$%d:$O$%d,"SIM")+%s!$L$%d' % (OR, O1, ONC, OR, O1, ONC, OR, OADM),
         note="Usa o maior entre orçado e tendência pelo avanço físico.")
    vrow(18, "terreno", "Terreno / permuta", "=CustoTerreno")
    vrow(19, "outras_com", "Outras despesas comerciais (marketing, stand...)",
         '=MAX(0,SUMIFS(%s!$L$%d:$L$%d,%s!$O$%d:$O$%d,"NÃO")-SUMIFS(D_VALOR,D_CAT,%s)-SUMIFS(D_VALOR,D_CAT,%s))' % (OR, O1, O2, OR, O1, O2, CAT_COMISSAO, CAT_IMPOSTO),
         note="Etapas 'NÃO custo de obra', sem comissões e impostos (já deduzidos acima).")
    vrow(20, "custo_tot", "CUSTO TOTAL PROJETADO", "=C17+C18+C19", strong=True)
    vi.merge_range("B22:D22", "RESULTADO", f_sec)
    vrow(23, "lucro", "LUCRO PROJETADO", "=C12-C20", strong=True)
    vrow(24, "margem", "Margem sobre o VGV", "=IFERROR(C23/C9,0)", PCT)
    vrow(25, "roi", "ROI sobre o custo total", "=IFERROR(C23/C20,0)", PCT)
    vrow(26, "cm2c", "Custo da obra por m² construído (projetado)", "=IFERROR(C17/AreaConstruida,0)")
    vrow(27, "cm2p", "Custo total por m² privativo (projetado)", "=IFERROR(C20/AreaPrivativa,0)")
    vrow(28, "pm2_tab", "Preço médio de tabela (R$/m² privativo)", '=IFERROR(C6/SUMIFS(V_AREA,V_STATUS,"<>PERMUTA"),0)')
    vrow(29, "pm2_vend", "Preço médio realizado nas vendas (R$/m²)", '=IFERROR(C7/SUMIFS(V_AREA,V_STATUS,"VENDIDA"),0)')
    vrow(30, "be_vgv", "Ponto de equilíbrio — VGV mínimo", "=IFERROR(C20/(1-IFERROR(C10/C9,ComissaoPadrao)-ImpostoVendas),0)",
         note="VGV que zera o lucro com as mesmas despesas de venda.")
    vrow(31, "be_m2", "Ponto de equilíbrio — preço mínimo por m²", '=IFERROR(C30/SUMIFS(V_AREA,V_STATUS,"<>PERMUTA"),0)')
    vrow(32, "expo", "Exposição máxima de caixa (necessidade de aporte)", "=MAX(0,-MIN(%s!$C$31:$%s$31))" % (FX, CL),
         note="Pior saldo projetado acumulado do FLUXO DE CAIXA.")
    vrow(33, "pct_vend", "% do VGV já vendido", "=IFERROR(C7/C9,0)", PCT)
    vrow(34, "pct_fis", "Avanço físico da obra", "=%s!$J$%d" % (OR, OSUB), PCT)

    vi.merge_range("F5:M5", "LUCRO PROJETADO POR SÓCIO", f_sec)
    vi.write_row("F6", ["SÓCIO", "PARTIC.", "APORTADO", "COTA DO LUCRO", "RETORNO TOTAL", "ROI S/ APORTE"], f_hdr)
    for k, sr in enumerate(SOC_ROWS):
        i = 6 + k
        R = str(i + 1)
        vi.write_formula(i, 5, '=IF(CONFIG!$B${0}="","",CONFIG!$B${0})'.format(sr), f_au(bold=True))
        vi.write_formula(i, 6, '=IF(F{0}="","",CONFIG!$C${1})'.format(R, sr), f_au(PCT))
        vi.write_formula(i, 7, '=IF(F{0}="","",SUMIFS(A_VALOR,A_SOCIO,F{0}))'.format(R), f_au(MOEDA))
        vi.write_formula(i, 8, '=IF(F{0}="","",G{0}*$C$23)'.format(R), f_au(MOEDA))
        vi.write_formula(i, 9, '=IF(F{0}="","",H{0}+I{0})'.format(R), f_au(MOEDA))
        vi.write_formula(i, 10, '=IF(F{0}="","",IFERROR(I{0}/H{0},0))'.format(R), f_au(PCT))
    tr = 6 + N_SOCIOS
    vi.write(tr, 5, "TOTAL", f_tot_l)
    vi.write_formula(tr, 6, "=SUM(G7:G%d)" % tr, f_tot_p)
    for col in (7, 8, 9):
        cl = xl_col_to_name(col)
        vi.write_formula(tr, col, "=SUM(%s7:%s%d)" % (cl, cl, tr), f_tot)

    # sensibilidade
    sr0 = tr + 3
    vi.merge_range(sr0 - 1, 5, sr0 - 1, 12, "SENSIBILIDADE DO LUCRO — variação de PREÇO (linhas) × variação de CUSTO DA OBRA (colunas)", f_sec)
    custos = [-0.10, -0.05, 0.0, 0.05, 0.10, 0.15, 0.20]
    precos = [-0.15, -0.10, -0.05, 0.0, 0.05, 0.10, 0.15]
    vi.write(sr0, 5, "Preço \\ Custo", f_hdr)
    for j, cv in enumerate(custos):
        vi.write(sr0, 6 + j, cv, fm(bold=True, font_color="white", bg_color=NAVY2, align="center", border=1, border_color=LINE, num_format='+0%;-0%;0%'))
    for i, pv in enumerate(precos):
        rr = sr0 + 1 + i
        vi.write(rr, 5, pv, fm(bold=True, font_color="white", bg_color=NAVY2, align="center", border=1, border_color=LINE, num_format='+0%;-0%;0%'))
        for j in range(len(custos)):
            pc = xl_rowcol_to_cell(rr, 5, col_abs=True)
            cc = xl_rowcol_to_cell(sr0, 6 + j, row_abs=True)
            fml = "=$C$9*(1+{p})*(1-IFERROR($C$10/$C$9,ComissaoPadrao)-ImpostoVendas)-($C$17*(1+{c})+$C$18+$C$19)".format(p=pc, c=cc)
            base = (i == 3 and j == 2)
            vi.write_formula(rr, 6 + j, fml, fm(border=1, border_color=LINE, num_format=MOEDA0, bold=base, bg_color=BRAND_L if base else "#FFFFFF"))
    sens_area = "G%d:M%d" % (sr0 + 2, sr0 + 1 + len(precos))
    vi.conditional_format(sens_area, {"type": "3_color_scale", "min_color": "#E8B4AE", "mid_type": "num", "mid_value": 0,
                                      "mid_color": "#FFFFFF", "max_color": "#C9D6C0"})
    VI["sens_r0"] = sr0
    # cenarios simples p/ grafico
    cr0 = sr0 + 2 + len(precos) + 1
    vi.merge_range(cr0 - 1, 5, cr0 - 1, 12, "CENÁRIOS DE PREÇO (custo projetado mantido)", f_sec)
    vi.write_row(cr0, 5, ["CENÁRIO", "PREÇO", "VGV", "LUCRO", "MARGEM", "LUCRO/SÓCIO (média)"], f_hdr)
    for i, (nm_, pv) in enumerate([("Pessimista", -0.10), ("Base", 0.0), ("Otimista", 0.10)]):
        rr = cr0 + 1 + i
        R = str(rr + 1)
        vi.write(rr, 5, nm_, f_lbl_b)
        vi.write(rr, 6, pv, f_au('+0%;-0%;0%'))
        vi.write_formula(rr, 7, "=$C$9*(1+G%s)" % R, f_au(MOEDA0))
        vi.write_formula(rr, 8, "=H{0}*(1-IFERROR($C$10/$C$9,ComissaoPadrao)-ImpostoVendas)-$C$20".format(R), f_au(MOEDA0))
        vi.write_formula(rr, 9, "=IFERROR(I{0}/H{0},0)".format(R), f_au(PCT))
        vi.write_formula(rr, 10, '=IFERROR(I{0}/COUNTIF(Socios,"?*"),0)'.format(R), f_au(MOEDA0))
    VI["cen_r0"] = cr0
    vi.conditional_format("C23:C23", {"type": "cell", "criteria": "<", "value": 0, "format": fm(font_color="#FFFFFF", bg_color=WINE, bold=True)})

    # =========================================================================
    # ACERTO MENSAL
    # =========================================================================
    a = ws["ACERTO MENSAL"]
    a.set_tab_color(BRAND_D)
    a.set_column("A:A", 2)
    a.set_column("B:B", 36)
    a.set_column("C:C", 12)
    a.set_column("D:D", 14)
    a.set_column("E:E", 22)
    a.set_column("F:F", 20)
    a.set_column("G:G", 14)
    a.set_column("H:H", 16)
    a.set_column("I:I", 17)
    a.set_column("J:J", 2)
    a.set_column("K:L", 6, None, {"hidden": True})
    a.set_row(1, 32)
    a.set_row(0, 22)
    a.merge_range("B1:I1", "HLM CONSTRUÇÕES  ·  ACERTO DE CONTAS COM OS SÓCIOS", fm(font_name=F_MONO, font_size=8, bold=True, font_color=BRAND, valign="bottom"))
    a.merge_range("B2:I2", "", f_title)
    a.write_formula("B2", '="ACERTO DE CONTAS  —  "&UPPER(NomeEmp)', f_title)
    a.merge_range("B3:I3", "", f_sub)
    a.write_formula("B3", '=Construtora&"  |  "&Conta&"  |  Taxa de administração: "&ROUND(TaxaAdm*100,2)&"%"', f_sub)
    a.set_row(4, 26)
    a.write("B5", "MÊS DE REFERÊNCIA  ▶", fm(bold=True, font_size=12, align="right", font_color=NAVY))
    f_sel = fm(bold=True, font_size=14, bg_color=BRAND_L, border=2, border_color=BRAND, align="center", num_format="mmmm/yyyy", font_name=F_TIT)
    a.merge_range("C5:D5", None, f_sel)
    if exemplo:
        a.write_datetime("C5", dt.datetime(2026, 9, 1), f_sel)  # HLM ONE: último mês com os dois caixas completos
    else:
        a.write_formula("C5", "=FluxoIni", f_sel)
    wb.define_name("SelMes", "=%s!$C$5" % AC)
    a.data_validation("C5", {"validate": "list", "source": "=Meses", "input_title": "Mês do acerto",
                             "input_message": "Escolha o mês. Todo o relatório é recalculado."})
    a.write("E5", "Emitido em:", fm(align="right", italic=True, font_color=WARM_D))
    a.write_formula("F5", "=TODAY()", fm(num_format=DATA, italic=True, font_color=WARM_D, align="left"))
    a.write_formula("K5", "=IFERROR(MATCH(SelMes,Meses,0),0)", fm(font_color=WARM))  # coluna do mes
    a.merge_range("B6:I6", "", f_note)
    if macro:
        a.write("B6", "Escolha o mês na célula laranja e clique em  BAIXAR ACERTO (PDF)  para gerar o arquivo e enviar aos sócios.", f_note)
        a.insert_button("G4", {"macro": "BaixarAcertoPDF", "caption": "⬇  BAIXAR ACERTO (PDF)", "width": 250, "height": 34,
                               "x_offset": 4, "y_offset": 4})
    else:
        a.write("B6", "Para baixar: Excel → Arquivo › Exportar › PDF (aba atual)  |  Google Planilhas → Arquivo › Fazer download › PDF › 'Página atual'. "
                      "A versão .xlsm tem o botão BAIXAR ACERTO (PDF).", f_note)
    a.write_formula("K6", "=IF(K5=0,1,0)", fm(font_color=WARM))

    def acol(fx_row, offset=0):
        """valor do FLUXO na coluna do mes selecionado (offset em meses)."""
        if offset == 0:
            return "IFERROR(INDEX(%s!$C$%d:$%s$%d,1,$K$5),0)" % (FX, fx_row, CL, fx_row, )
        return "IFERROR(INDEX(%s!$C$%d:$%s$%d,1,$K$5%+d),0)" % (FX, fx_row, CL, fx_row, offset)

    # 1. custos por categoria
    a.merge_range("B8:I8", "1.  CUSTOS PAGOS NO MÊS POR CATEGORIA", f_sec)
    a.write_row("B9", ["CATEGORIA", "", "OFICIAL (c/ nota)", "ESPECIAL (s/ nota)", "TOTAL", "% DO MÊS"], f_hdr)
    a.merge_range("B9:C9", "CATEGORIA", f_hdr)
    a.write_blank("H9", None, f_hdr)
    a.write_blank("I9", None, f_hdr)
    ac_r1 = 10
    for k in range(N_CAT):
        r = ac_r1 + k
        R = str(r)
        cr = 6 + k
        a.merge_range("B%s:C%s" % (R, R), None, f_lbl)
        a.write_formula("B" + R, '=IF(CONFIG!$F${0}="","",CONFIG!$F${0})'.format(cr), f_lbl)
        a.write_formula("D" + R, '=IF($B{0}="","",SUMIFS(D_VALOR,D_MES,SelMes,D_TIPO,"OFICIAL",D_CAT,$B{0},D_STATUS,"<>A PAGAR"))'.format(R), f_au(MOEDA))
        a.write_formula("E" + R, '=IF($B{0}="","",SUMIFS(D_VALOR,D_MES,SelMes,D_TIPO,"ESPECIAL",D_CAT,$B{0},D_STATUS,"<>A PAGAR"))'.format(R), f_au(MOEDA))
        a.write_formula("F" + R, '=IF($B{0}="","",SUMIFS(D_VALOR,D_MES,SelMes,D_CAT,$B{0},D_STATUS,"<>A PAGAR"))'.format(R), f_au(MOEDA, bold=True))
        a.write_formula("G" + R, '=IF($B{0}="","",IFERROR(F{0}/$F${1},0))'.format(R, ac_r1 + N_CAT + 2), f_au(PCT))
        a.write_blank("H" + R, None, None)
    rnc = ac_r1 + N_CAT          # 30 nao classificado
    rfo = rnc + 1                # 31 folha
    rsub = rfo + 1               # 32 subtotal
    radm = rsub + 1              # 33 adm
    rtot = radm + 1              # 34 total
    a.merge_range("B%d:C%d" % (rnc, rnc), "Despesas sem categoria/tipo válido (corrigir)", fm(border=1, border_color=LINE, indent=1, italic=True, font_color=RED_FG))
    a.write_formula("D%d" % rnc, '=SUMIFS(D_VALOR,D_MES,SelMes,D_TIPO,"OFICIAL",D_STATUS,"<>A PAGAR")-SUM(D%d:D%d)' % (ac_r1, rnc - 1), f_au(MOEDA))
    a.write_formula("E%d" % rnc, '=SUMIFS(D_VALOR,D_MES,SelMes,D_TIPO,"ESPECIAL",D_STATUS,"<>A PAGAR")-SUM(E%d:E%d)' % (ac_r1, rnc - 1), f_au(MOEDA))
    a.write_formula("F%d" % rnc, '=SUMIFS(D_VALOR,D_MES,SelMes,D_STATUS,"<>A PAGAR")-SUM(F%d:F%d)' % (ac_r1, rnc - 1), f_au(MOEDA, bold=True))
    a.write_formula("G%d" % rnc, "=IFERROR(F%d/$F$%d,0)" % (rnc, rsub), f_au(PCT))
    a.merge_range("B%d:C%d" % (rfo, rfo), "Mão de obra — folha de pagamento", f_lbl_b)
    a.write_formula("D%d" % rfo, "=" + acol(20), f_au(MOEDA))
    a.write_formula("E%d" % rfo, "=" + acol(21), f_au(MOEDA))
    a.write_formula("F%d" % rfo, "=D%d+E%d" % (rfo, rfo), f_au(MOEDA, bold=True))
    a.write_formula("G%d" % rfo, "=IFERROR(F%d/$F$%d,0)" % (rfo, rsub), f_au(PCT))
    a.merge_range("B%d:C%d" % (rsub, rsub), "SUBTOTAL — CUSTO DIRETO DO MÊS", f_tot_l)
    for cl in "DEF":
        a.write_formula("%s%d" % (cl, rsub), "=SUM(%s%d:%s%d)" % (cl, ac_r1, cl, rfo), f_tot)
    a.write_formula("G%d" % rsub, "=IF(F%d=0,0,1)" % rsub, f_tot_p)
    a.merge_range("B%d:C%d" % (radm, radm), None, f_lbl)
    a.write_formula("B%d" % radm, '="Taxa de administração ("&ROUND(TaxaAdm*100,2)&"%)"', f_lbl)
    a.write_formula("F%d" % radm, "=" + acol(23), f_au(MOEDA, bold=True))
    a.write_blank("D%d" % radm, None, f_au())
    a.write_blank("E%d" % radm, None, f_au())
    a.merge_range("B%d:E%d" % (rtot, rtot), "TOTAL DO MÊS A RATEAR (custo direto + taxa de adm)", f_big_l)
    a.write_formula("F%d" % rtot, "=F%d+F%d" % (rsub, radm), f_big)
    a.write_formula("H%d" % rsub, '=IF(ABS(F{0}-{1})<0.01,"✔ confere com o fluxo","✖ difere do fluxo")'.format(rsub, acol(22)),
                    fm(italic=True, font_size=9, font_color=WARM_D))
    a.write_blank("G%d" % rtot, None, f_big)

    # 2. resumo de caixa
    s2 = rtot + 2
    a.merge_range("B%d:I%d" % (s2, s2), "2.  MOVIMENTO DE CAIXA DO MÊS (conta da obra)", f_sec)
    caixa = [
        ("Saldo inicial (acumulado até o mês anterior)", "=IF($K$5<=1,0,%s)" % acol(27, -1)),
        ("(+) Aportes dos sócios", "=%s-%s" % (acol(15), acol(14))),
        ("(+) Recebimentos de vendas", "=" + acol(14)),
        ("(−) Custo direto pago", "=F%d" % rsub),
        ("(−) Taxa de administração", "=F%d" % radm),
    ]
    rr = s2 + 1
    for lbl, fml in caixa:
        a.merge_range("B%d:E%d" % (rr, rr), lbl, f_lbl)
        a.write_formula("F%d" % rr, fml, f_au(MOEDA))
        rr += 1
    a.merge_range("B%d:E%d" % (rr, rr), "(=) SALDO FINAL DE CAIXA", f_tot_l)
    a.write_formula("F%d" % rr, "=F{0}+F{1}+F{2}-F{3}-F{4}".format(s2 + 1, s2 + 2, s2 + 3, s2 + 4, s2 + 5), f_tot)
    a.write_formula("G%d" % rr, '=IF(ABS(F{0}-{1})<0.01,"✔ confere","✖ verificar")'.format(rr, acol(27)), fm(italic=True, font_size=9, font_color=WARM_D))
    r_saldo = rr
    rr += 1
    a.merge_range("B%d:E%d" % (rr, rr), "Contas a pagar com vencimento no mês seguinte (c/ adm)", f_lbl)
    a.write_formula("F%d" % rr, "=" + acol(29, 1), f_au(MOEDA))
    rr += 1
    a.merge_range("B%d:E%d" % (rr, rr), "Recebimentos previstos no mês seguinte", f_lbl)
    a.write_formula("F%d" % rr, "=" + acol(30, 1), f_au(MOEDA))
    a.conditional_format("F%d" % r_saldo, {"type": "cell", "criteria": "<", "value": -0.005, "format": fm(font_color="#FFFFFF", bg_color=WINE, bold=True, num_format=MOEDA)})

    # 3. rateio por socio
    s3 = rr + 2
    a.merge_range("B%d:I%d" % (s3, s3), "3.  RATEIO E POSIÇÃO DE CADA SÓCIO", f_sec)
    a.set_row(s3, 32)
    a.write_row(s3, 1, ["SÓCIO", "PARTIC.", "COTA DOS CUSTOS DO MÊS", "APORTES NO MÊS", "COTA DAS RECEITAS DO MÊS",
                        "RESULTADO DO MÊS", "POSIÇÃO ACUMULADA", "SITUAÇÃO"], f_hdr)
    soc_a1 = s3 + 2
    for k, sr in enumerate(SOC_ROWS):
        r = soc_a1 + k
        R = str(r)
        a.write_formula("B" + R, '=IF(CONFIG!$B${0}="","",CONFIG!$B${0})'.format(sr), f_lbl_b)
        a.write_formula("C" + R, '=IF(B{0}="","",CONFIG!$C${1})'.format(R, sr), f_au(PCT))
        a.write_formula("D" + R, '=IF(B{0}="","",$F${1}*C{0})'.format(R, rtot), f_au(MOEDA))
        a.write_formula("E" + R, '=IF(B{0}="","",SUMIFS(A_VALOR,A_SOCIO,B{0},A_MES,SelMes))'.format(R), f_au(MOEDA))
        a.write_formula("F" + R, '=IF(B{0}="","",$F${1}*C{0})'.format(R, s2 + 3), f_au(MOEDA))
        a.write_formula("G" + R, '=IF(B{0}="","",E{0}+F{0}-D{0})'.format(R), f_au(MOEDA, bold=True))
        a.write_formula("H" + R, ('=IF(B{0}="","",SUMIFS(A_VALOR,A_SOCIO,B{0},A_MES,">="&FluxoIni,A_MES,"<="&SelMes)'
                                  '+C{0}*({1}-{2}))').format(R, acol(36), acol(34)), f_au(MOEDA, bold=True))
        a.write_formula("I" + R, '=IF(B{0}="","",IF(H{0}<-0.005,"DEVE APORTAR",IF(H{0}>0.005,"CRÉDITO","EM DIA")))'.format(R),
                        f_au(align="center", bold=True))
    soc_a2 = soc_a1 + N_SOCIOS - 1
    rts = soc_a2 + 1
    a.write("B%d" % rts, "TOTAL", f_tot_l)
    a.write_formula("C%d" % rts, "=SUM(C%d:C%d)" % (soc_a1, soc_a2), f_tot_p)
    for cl in "DEFGH":
        a.write_formula("%s%d" % (cl, rts), "=SUM(%s%d:%s%d)" % (cl, soc_a1, cl, soc_a2), f_tot)
    a.write_formula("I%d" % rts, '=IF(ROUND(C%d,4)<>1,"⚠ PARTIC. ≠ 100%%","")' % rts, fm(bold=True, font_color=RED_FG, bg_color="#ECE7E0", border=1, border_color=LINE))
    a.merge_range("B%d:I%d" % (rts + 1, rts + 2),
                  "Cota dos custos = total do mês × participação.  Resultado do mês = aportes + cota das receitas − cota dos custos.  "
                  "Posição acumulada = tudo o que o sócio aportou + sua cota das receitas − sua cota de todos os custos, desde o início até este mês. "
                  "Negativo = DEVE APORTAR esse valor;  positivo = CRÉDITO (aportou acima da sua parte).", f_note)
    a.set_row(rts, 18)
    a.conditional_format("I%d:I%d" % (soc_a1, soc_a2), {"type": "cell", "criteria": "==", "value": '"DEVE APORTAR"', "format": fm(bg_color=RED_BG, font_color=RED_FG, bold=True)})
    a.conditional_format("I%d:I%d" % (soc_a1, soc_a2), {"type": "cell", "criteria": "==", "value": '"CRÉDITO"', "format": fm(bg_color=GREEN_BG, font_color=GREEN_FG, bold=True)})
    a.conditional_format("I%d:I%d" % (soc_a1, soc_a2), {"type": "cell", "criteria": "==", "value": '"EM DIA"', "format": fm(bg_color=CREAM, font_color=INK, bold=True)})

    # 4. detalhamento despesas
    s4 = rts + 4
    a.merge_range("B%d:I%d" % (s4, s4), "", f_sec)
    a.write_formula("B%d" % s4, '="4.  DESPESAS PAGAS NO MÊS — "&COUNTIFS(D_MES,SelMes,D_VALOR,"<>",D_STATUS,"<>A PAGAR")&" lançamento(s)"', f_sec)
    a.write_row(s4, 1, ["FORNECEDOR / DESCRIÇÃO", "DATA", "TIPO", "CATEGORIA", "ETAPA", "NF / RECIBO", "FORMA", "VALOR"], f_hdr)
    d1 = s4 + 2
    for k in range(DET_LINHAS):
        r = d1 + k
        R = str(r)
        a.write_formula("K" + R, '=IFERROR(MATCH(%d,D_ORD,0),"")' % (k + 1), fm(font_color=WARM, font_size=8))

        def g(rng_name, R=R):
            return 'IF($K{0}="","",INDEX({1},$K{0}))'.format(R, rng_name)

        a.write_formula("B" + R, '=IF($K{0}="","",INDEX(D_FORN,$K{0})&IF(AND(INDEX(D_FORN,$K{0})<>"",INDEX(D_DESC,$K{0})<>"")," — ","")&INDEX(D_DESC,$K{0}))'.format(R),
                        fm(border=1, border_color=LINE, font_size=9, indent=1, shrink=True))
        a.write_formula("C" + R, "=" + g("D_DATA"), fm(border=1, border_color=LINE, font_size=9, num_format=DATA, align="center"))
        a.write_formula("D" + R, "=" + g("D_TIPO"), fm(border=1, border_color=LINE, font_size=9, align="center"))
        a.write_formula("E" + R, "=" + g("D_CAT"), fm(border=1, border_color=LINE, font_size=9, shrink=True))
        a.write_formula("F" + R, "=" + g("D_ETAPA"), fm(border=1, border_color=LINE, font_size=9, shrink=True))
        a.write_formula("G" + R, '=IF($K{0}="","",INDEX(D_NF,$K{0})&"")'.format(R), fm(border=1, border_color=LINE, font_size=9, align="center"))
        a.write_formula("H" + R, '=IF($K{0}="","",INDEX(D_FORMA,$K{0})&"")'.format(R), fm(border=1, border_color=LINE, font_size=9, align="center"))
        a.write_formula("I" + R, "=" + g("D_VALOR"), fm(border=1, border_color=LINE, font_size=9, num_format=MOEDA))
    d2 = d1 + DET_LINHAS - 1
    wb.define_name("AcertoDetalhe", "=%s!$K$%d:$K$%d" % (AC, d1, d2))
    a.write("B%d" % (d2 + 1), "Total detalhado", f_tot_l)
    a.write_formula("I%d" % (d2 + 1), "=SUM(I%d:I%d)" % (d1, d2), f_tot)
    a.merge_range("C%d:H%d" % (d2 + 1, d2 + 1), "", f_tot_l)
    a.write_formula("C%d" % (d2 + 1), '=IF(COUNTIFS(D_MES,SelMes,D_VALOR,"<>",D_STATUS,"<>A PAGAR")>%d,"⚠ Há mais de %d lançamentos no mês — veja a aba DESPESAS (filtro por MÊS).","")' % (DET_LINHAS, DET_LINHAS),
                    fm(bold=True, font_color=RED_FG, bg_color="#ECE7E0", border=1, border_color=LINE))

    # 5. folha
    s5 = d2 + 3
    a.merge_range("B%d:I%d" % (s5, s5), "", f_sec)
    a.write_formula("B%d" % s5, '="5.  FOLHA DE PAGAMENTO DO MÊS — "&COUNTIFS(F_MES,SelMes,F_TOTAL,"<>")&" pagamento(s)"', f_sec)
    a.write_row(s5, 1, ["COLABORADOR", "DATA", "CAIXA", "QUINZENA", "ETAPA", "DIAS × DIÁRIA", "EXTRAS − DESC.", "TOTAL"], f_hdr)
    f1 = s5 + 2
    for k in range(FOL_LINHAS):
        r = f1 + k
        R = str(r)
        a.write_formula("K" + R, '=IFERROR(MATCH(%d,F_ORD,0),"")' % (k + 1), fm(font_color=WARM, font_size=8))
        a.write_formula("B" + R, '=IF($K{0}="","",INDEX(F_NOME,$K{0})&IF(INDEX(F_FUNCAO,$K{0})<>""," ("&INDEX(F_FUNCAO,$K{0})&")",""))'.format(R),
                        fm(border=1, border_color=LINE, font_size=9, indent=1, shrink=True))
        a.write_formula("C" + R, '=IF($K{0}="","",INDEX(F_DATA,$K{0}))'.format(R), fm(border=1, border_color=LINE, font_size=9, num_format=DATA, align="center"))
        a.write_formula("D" + R, '=IF($K{0}="","",IF(INDEX(F_CAIXA,$K{0})="OFICIAL","OFICIAL","ESPECIAL"))'.format(R), fm(border=1, border_color=LINE, font_size=9, align="center"))
        a.write_formula("E" + R, '=IF($K{0}="","",INDEX(F_QUINZ,$K{0})&"")'.format(R), fm(border=1, border_color=LINE, font_size=9, shrink=True))
        a.write_formula("F" + R, '=IF($K{0}="","",INDEX(F_ETAPA,$K{0})&"")'.format(R), fm(border=1, border_color=LINE, font_size=9, shrink=True))
        a.write_formula("G" + R, '=IF($K{0}="","",N(INDEX(F_DIAS,$K{0}))*INDEX(F_DIARIA,$K{0}))'.format(R), fm(border=1, border_color=LINE, font_size=9, num_format=MOEDA))
        a.write_formula("H" + R, '=IF($K{0}="","",N(INDEX(F_EXTRA,$K{0}))-N(INDEX(F_DESC,$K{0})))'.format(R), fm(border=1, border_color=LINE, font_size=9, num_format=MOEDA))
        a.write_formula("I" + R, '=IF($K{0}="","",INDEX(F_TOTAL,$K{0}))'.format(R), fm(border=1, border_color=LINE, font_size=9, num_format=MOEDA))
    f2 = f1 + FOL_LINHAS - 1
    wb.define_name("AcertoFolha", "=%s!$K$%d:$K$%d" % (AC, f1, f2))
    a.write("B%d" % (f2 + 1), "Total da folha", f_tot_l)
    a.merge_range("C%d:H%d" % (f2 + 1, f2 + 1), "", f_tot_l)
    a.write_formula("I%d" % (f2 + 1), "=SUM(I%d:I%d)" % (f1, f2), f_tot)

    # assinaturas
    s6 = f2 + 4
    a.write("B%d" % s6, "______________________________", fm(align="center"))
    a.write("E%d" % s6, "______________________________", fm(align="center"))
    a.merge_range("G%d:I%d" % (s6, s6), "______________________________", fm(align="center"))
    a.write_formula("B%d" % (s6 + 1), '="Responsável — "&Construtora', fm(align="center", font_size=9, font_color=WARM_D))
    a.write("E%d" % (s6 + 1), "Sócio", fm(align="center", font_size=9, font_color=WARM_D))
    a.merge_range("G%d:I%d" % (s6 + 1, s6 + 1), "Sócio", fm(align="center", font_size=9, font_color=WARM_D))

    a.print_area("B1:I%d" % (s6 + 1))
    a.set_portrait()
    a.set_paper(9)
    a.fit_to_pages(1, 0)
    a.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.6)
    a.set_footer('&L&8&A&C&8Página &P de &N&R&8Gerado pela planilha de gestão HLM')
    a.center_horizontally()
    a.freeze_panes(6, 0)

    # =========================================================================
    # CALC (oculta) — dados para graficos
    # =========================================================================
    k_ = ws["CALC"]
    k_.hide()
    k_.write_row("A1", ["CATEGORIA", "TOTAL PAGO"])
    for i in range(N_CAT):
        r = 2 + i
        k_.write_formula(r - 1, 0, '=IF(CONFIG!$F${0}="","",CONFIG!$F${0})'.format(6 + i))
        k_.write_formula(r - 1, 1, '=IF(A{0}="",0,SUMIFS(D_VALOR,D_CAT,A{0},D_STATUS,"<>A PAGAR",D_MES,">="&FluxoIni,D_MES,"<="&FluxoFim))'.format(r))
    cfol = 2 + N_CAT
    k_.write(cfol - 1, 0, "Mão de obra (folha)")
    k_.write_formula(cfol - 1, 1, "=%s!$B$20+%s!$B$21" % (FX, FX))
    k_.write(cfol, 0, "Taxa de administração")
    k_.write_formula(cfol, 1, "=%s!$B$23" % FX)
    k_.write(cfol + 1, 0, "Sem categoria (corrigir)")
    k_.write_formula(cfol + 1, 1, "=MAX(0,%s!$B$17+%s!$B$18+%s!$B$19-SUM(B2:B%d))" % (FX, FX, FX, cfol - 1))
    CALC_CAT = (2, cfol + 2)
    # socios
    k_.write_row("D1", ["SÓCIO", "APORTADO", "COTA LÍQUIDA"])
    for k, sr in enumerate(SOC_ROWS):
        r = 2 + k
        k_.write_formula(r - 1, 3, '=IF(CONFIG!$B${0}="","",CONFIG!$B${0})'.format(sr))
        k_.write_formula(r - 1, 4, "=IF(D{0}=\"\",0,APORTES!K{1})".format(r, 6 + k))
        k_.write_formula(r - 1, 5, "=IF(D{0}=\"\",0,APORTES!L{1})".format(r, 6 + k))
    # oficial x especial
    k_.write_row("H1", ["CAIXA", "VALOR"])
    k_.write("H2", "Oficial (com nota)")
    k_.write_formula("I2", "=%s!$B$17+%s!$B$20" % (FX, FX))
    k_.write("H3", "Especial (sem nota)")
    k_.write_formula("I3", "=%s!$B$18+%s!$B$21" % (FX, FX))
    k_.write("H4", "Sem tipo (corrigir)")
    k_.write_formula("I4", "=%s!$B$19" % FX)
    # status unidades
    k_.write_row("K1", ["STATUS", "UNIDADES", "VGV"])
    for i, st in enumerate(STATUS_UN):
        r = 2 + i
        k_.write(r - 1, 10, st)
        k_.write_formula(r - 1, 11, '=COUNTIF(V_STATUS,K%d)' % r)
        if st == "VENDIDA":
            k_.write_formula(r - 1, 12, '=SUMIFS(V_VALOR,V_STATUS,K%d)' % r)
        else:
            k_.write_formula(r - 1, 12, '=SUMIFS(V_TAB,V_STATUS,K%d)' % r)
    # series temporais que param no mes atual (graficos de linha)
    k_.write("A30", "MÊS")
    k_.write("A31", "Saldo acumulado (até hoje)")
    k_.write("A32", "Custo acumulado (até hoje)")
    for kk in range(N_MESES):
        cl = xl_col_to_name(2 + kk)
        k_.write_formula("%s30" % cl, "=%s!%s5" % (FX, cl))
        k_.write_formula("%s31" % cl, "=IF(%s30<=TODAY(),%s!%s27,NA())" % (cl, FX, cl))
        k_.write_formula("%s32" % cl, "=IF(%s30<=TODAY(),%s!%s34,NA())" % (cl, FX, cl))
    # etapas curtas p/ grafico
    k_.write_row("O1", ["ETAPA", "ORÇADO", "REALIZADO", "% FIN", "% FÍS"])
    for i in range(N_ETAPA):
        r = 2 + i
        k_.write_formula(r - 1, 14, "=%s!A%d" % (OR, O1 + i))
        k_.write_formula(r - 1, 15, "=%s!B%d" % (OR, O1 + i))
        k_.write_formula(r - 1, 16, "=%s!F%d" % (OR, O1 + i))
        k_.write_formula(r - 1, 17, "=%s!I%d" % (OR, O1 + i))
        k_.write_formula(r - 1, 18, "=N(%s!J%d)" % (OR, O1 + i))

    # =========================================================================
    # AUDITORIA
    # =========================================================================
    au = ws["AUDITORIA"]
    au.set_tab_color(WARM_D)
    au.set_column("A:A", 2)
    au.set_column("B:B", 5)
    au.set_column("C:C", 66)
    au.set_column("D:D", 18)
    au.set_column("E:E", 18)
    au.set_column("F:F", 14)
    au.set_row(1, 30)
    au.merge_range("B2:F2", "", f_title)
    au.write_formula("B2", nome_titulo("AUDITORIA E CONCILIAÇÃO"), f_title)
    au.merge_range("B3:F3", "Verificações automáticas de consistência. Tudo deve ficar OK antes de enviar o acerto aos sócios.", f_sub)
    au.write_row("B5", ["Nº", "VERIFICAÇÃO", "ONDE CORRIGIR", "OCORRÊNCIAS / DIF.", "STATUS"], f_hdr)
    checks = [
        ("Lançamentos de despesa com erro (data, tipo, categoria, etapa, status, vencimento)", "DESPESAS (col. N)", '=COUNTIFS(D_CHECK,"?*",D_CHECK,"<>OK")', "0"),
        ("Lançamentos da folha com erro (data, colaborador, diária, etapa, caixa)", "FOLHA PAG (col. O)", '=COUNTIFS(F_CHECK,"?*",F_CHECK,"<>OK")', "0"),
        ("Aportes com erro (data, sócio, valor)", "APORTES (col. G)", '=COUNTIFS(A_CHECK,"?*",A_CHECK,"<>OK")', "0"),
        ("Unidades com cadastro/venda inconsistente", "VENDAS (col. R)", '=COUNTIFS(V_CHECK,"?*",V_CHECK,"<>OK")', "0"),
        ("Recebimentos com erro (exceto vencidos)", "RECEBIMENTOS (col. J)", '=COUNTIFS(R_CHECK,"?*",R_CHECK,"<>OK",R_CHECK,"<>VENCIDO (inadimplência?)")', "0"),
        ("Soma das participações dos sócios ≠ 100%", "CONFIG", "=IF(ROUND(PartTotal,4)=1,0,1)", "0"),
        ("Conciliação: despesas pagas lançadas × fluxo de caixa (diferença R$)", "DESPESAS — datas fora do período",
         '=ROUND(SUMIFS(D_VALOR,D_STATUS,"<>A PAGAR")-(%s!$B$17+%s!$B$18+%s!$B$19),2)' % (FX, FX, FX), MOEDA),
        ("Conciliação: folha lançada × fluxo de caixa (diferença R$)", "FOLHA PAG — datas fora do período",
         "=ROUND(SUM(F_TOTAL)-(%s!$B$20+%s!$B$21),2)" % (FX, FX), MOEDA),
        ("Conciliação: aportes lançados × fluxo de caixa (diferença R$)", "APORTES — datas fora do período",
         "=ROUND(SUM(A_VALOR)-SUM(%s!$B$7:$B$13),2)" % FX, MOEDA),
        ("Conciliação: recebimentos (RECEBIDO) × fluxo de caixa (diferença R$)", "RECEBIMENTOS — datas fora do período",
         '=ROUND(SUMIFS(R_VALOR,R_STATUS,"RECEBIDO")-%s!$B$14,2)' % FX, MOEDA),
        ("Conciliação: realizado do ORÇAMENTO × custo direto do fluxo (diferença R$)", "ORÇAMENTO / datas",
         "=ROUND(%s!$F$%d-%s!$B$22,2)" % (OR, OSUB, FX), MOEDA),
        ("Valores sem sócio / sem tipo / sem etapa (linhas 'corrigir')", "FLUXO l.13 e 19 / ORÇAMENTO l.%d" % ONC,
         "=ROUND(ABS(%s!$B$13)+ABS(%s!$B$19)+ABS(%s!$F$%d)+ABS(%s!$G$%d),2)" % (FX, FX, OR, ONC, OR, ONC), MOEDA),
        ("Etapas com orçamento ESTOURADO ou SEM ORÇAMENTO", "ORÇAMENTO (col. N)",
         '=COUNTIF(%s!$N$%d:$N$%d,"ESTOURADO")+COUNTIF(%s!$N$%d:$N$%d,"SEM ORÇAMENTO")' % (OR, O1, ONC, OR, O1, ONC), "0"),
        ("Contas a pagar VENCIDAS (quantidade)", "DESPESAS — STATUS / VENCIMENTO", '=COUNTIFS(D_STATUS,"A PAGAR",D_VENC,"<"&TODAY())', "0"),
        ("Parcelas a receber VENCIDAS (quantidade)", "RECEBIMENTOS", '=COUNTIFS(R_STATUS,"A RECEBER",R_DATA,"<"&TODAY())', "0"),
        ("Meses com saldo de caixa acumulado negativo", "FLUXO DE CAIXA l.27", "=SUMPRODUCT(--(%s!$C$27:$%s$27<-0.005))" % (FX, CL), "0"),
    ]
    for i, (txt, onde, fml, nf) in enumerate(checks):
        r = 5 + i
        au.write(r, 1, i + 1, f_au(align="center"))
        au.write(r, 2, txt, fm(border=1, border_color=LINE, indent=1, text_wrap=True))
        au.write(r, 3, onde, fm(border=1, border_color=LINE, font_size=9, italic=True, text_wrap=True))
        au.write_formula(r, 4, fml, f_au(nf if nf != "0" else "0", align="center"))
        au.write_formula(r, 5, '=IF(ABS(E%d)<0.005,"OK","CORRIGIR")' % (r + 1), f_au(align="center", bold=True))
    a1, a2 = 6, 5 + len(checks)
    au.write(a2, 2, "TOTAL DE VERIFICAÇÕES COM PENDÊNCIA", f_tot_l)
    au.write_formula(a2, 5, '=COUNTIF(F%d:F%d,"CORRIGIR")' % (a1, a2), fm(bold=True, bg_color="#ECE7E0", border=1, border_color=LINE, align="center"))
    wb.define_name("AudPend", "=AUDITORIA!$F$%d" % (a2 + 1))
    au.conditional_format("F%d:F%d" % (a1, a2), {"type": "cell", "criteria": "==", "value": '"OK"', "format": fm(bg_color=GREEN_BG, font_color=GREEN_FG, bold=True)})
    au.conditional_format("F%d:F%d" % (a1, a2), {"type": "cell", "criteria": "==", "value": '"CORRIGIR"', "format": fm(bg_color=RED_BG, font_color=RED_FG, bold=True)})

    # =========================================================================
    # PAINEL GERAL
    # =========================================================================
    p = ws["PAINEL GERAL"]
    p.set_tab_color(BRAND)
    p.set_column("A:A", 2)
    p.set_column("B:Q", 10.5)
    p.set_column("R:R", 2)
    p.set_row(1, 34)
    p.set_row(0, 22)
    p.merge_range("B1:Q1", "HLM CONSTRUÇÕES  ·  PAINEL DE GESTÃO DA OBRA", fm(font_name=F_MONO, font_size=8, bold=True, font_color=BRAND, valign="bottom"))
    p.merge_range("B2:Q2", "", fm(bold=True, font_size=20, font_color="white", bg_color=INK, indent=1, font_name=F_TIT, bottom=5, bottom_color=BRAND))
    p.write_formula("B2", '="PAINEL GERAL  —  "&UPPER(NomeEmp)', fm(bold=True, font_size=20, font_color="white", bg_color=INK, indent=1, font_name=F_TIT, bottom=5, bottom_color=BRAND))
    p.merge_range("B3:Q3", "", f_sub)
    p.write_formula("B3", '=Construtora&"  |  "&IF(Endereco="","",Endereco&"  |  ")&NumUnid&" unidades  |  Prazo de obra: "&Prazo&" meses  |  Pendências na auditoria: "&AudPend', f_sub)

    def tile(ws_, r, c0, label, fml, nf, sub_lbl, sub_fml, sub_nf, color):
        f_l = fm(bold=True, font_size=8, font_color=BRAND if color in (INK, BRAND) else CREAM, bg_color=INK, indent=1, font_name=F_MONO)
        f_v = fm(bold=True, font_size=18, font_color=color, bg_color="#FFFFFF", indent=1, num_format=nf, font_name=F_TIT,
                 left=1, right=1, left_color=LINE, right_color=LINE)
        f_sl = fm(font_size=8, font_color=WARM_D, bg_color=CREAM_L, indent=1, left=1, left_color=LINE, bottom=5, bottom_color=color,
                  font_name=F_MONO)
        f_sv = fm(font_size=9, bold=True, font_color=INK, bg_color=CREAM_L, align="right", num_format=sub_nf,
                  right=1, right_color=LINE, bottom=5, bottom_color=color)
        ws_.merge_range(r, c0, r, c0 + 3, label, f_l)
        ws_.merge_range(r + 1, c0, r + 1, c0 + 3, "", f_v)
        ws_.write_formula(r + 1, c0, fml, f_v)
        ws_.merge_range(r + 2, c0, r + 2, c0 + 1, sub_lbl, f_sl)
        ws_.merge_range(r + 2, c0 + 2, r + 2, c0 + 3, "", f_sv)
        ws_.write_formula(r + 2, c0 + 2, sub_fml, f_sv)
        ws_.set_row(r + 1, 30)

    p.merge_range("B5:Q5", "FINANCEIRO DA OBRA", f_sec)
    tile(p, 5, 1, "TOTAL APORTADO PELOS SÓCIOS", "=SUM(%s!$B$7:$B$13)" % FX, MOEDA0, "Sócios ativos", '=COUNTIF(Socios,"?*")', "0", INK)
    tile(p, 5, 5, "CUSTO REALIZADO (c/ adm)", "=%s!$B$24" % FX, MOEDA0, "Taxa de adm", "=%s!$B$23" % FX, MOEDA0, BRAND)
    tile(p, 5, 9, "SALDO EM CAIXA", "=%s!$B$27" % FX, MOEDA0, "Saldo projetado", "=%s!$B$31" % FX, MOEDA0, BRAND_D)
    tile(p, 5, 13, "CONTAS A PAGAR", '=SUMIFS(D_VALOR,D_STATUS,"A PAGAR")', MOEDA0, "Vencidas", '=SUMIFS(D_VALOR,D_STATUS,"A PAGAR",D_VENC,"<"&TODAY())', MOEDA0, WINE)
    p.merge_range("B10:Q10", "ORÇAMENTO E AVANÇO", f_sec)
    tile(p, 10, 1, "ORÇAMENTO TOTAL (c/ adm)", "=%s!$B$%d" % (OR, OTOT), MOEDA0, "% consumido", "=%s!$I$%d" % (OR, OTOT), PCT, INK)
    tile(p, 10, 5, "AVANÇO FÍSICO", "=%s!$J$%d" % (OR, OSUB), PCT, "Avanço financeiro", "=%s!$I$%d" % (OR, OSUB), PCT, WARM_D)
    tile(p, 10, 9, "CUSTO PROJETADO (EAC)", "=%s!$L$%d" % (OR, OTOT), MOEDA0, "Variação × orçado", "=%s!$M$%d" % (OR, OTOT), MOEDA0, "#3A3633")
    tile(p, 10, 13, "CUSTO / m² CONSTRUÍDO (proj.)", "=vb_cm2c", MOEDA0, "Realizado / m²", "=IFERROR(%s!$B$24/AreaConstruida,0)" % FX, MOEDA0, BRAND_D)
    p.merge_range("B15:Q15", "COMERCIAL E RESULTADO", f_sec)
    tile(p, 15, 1, "VGV TOTAL ESPERADO", "=vb_vgv", MOEDA0, "VGV vendido", "=vb_vgv_vend", MOEDA0, INK)
    tile(p, 15, 5, "UNIDADES VENDIDAS", '=COUNTIF(V_STATUS,"VENDIDA")&" de "&COUNTIFS(V_UNID,"?*",V_STATUS,"<>PERMUTA")', "@", "% do VGV vendido", "=vb_pct_vend", PCT, WARM_D)
    tile(p, 15, 9, "LUCRO PROJETADO", "=vb_lucro", MOEDA0, "Margem s/ VGV", "=vb_margem", PCT, BRAND_D)
    tile(p, 15, 13, "RECEBIDO DE CLIENTES", '=SUMIFS(R_VALOR,R_STATUS,"RECEBIDO")', MOEDA0, "A receber", '=SUMIFS(R_VALOR,R_STATUS,"A RECEBER")', MOEDA0, BRAND)

    # alertas
    p.merge_range("B20:Q20", "ALERTAS DE GESTÃO", f_sec)
    alerts = [
        ("Pendências de lançamento / conciliação (aba AUDITORIA)", "=AudPend", "0", "=IF(AudPend=0,\"OK\",\"ATENÇÃO\")"),
        ("Etapas com orçamento estourado", '=COUNTIF(%s!$N$%d:$N$%d,"ESTOURADO")' % (OR, O1, ONC), "0", None),
        ("Etapas em atenção (desvio físico-financeiro)", '=COUNTIF(%s!$N$%d:$N$%d,"ATENÇÃO")' % (OR, O1, ONC), "0", None),
        ("Contas a pagar vencidas (quantidade)", '=COUNTIFS(D_STATUS,"A PAGAR",D_VENC,"<"&TODAY())', "0", None),
        ("Parcelas de clientes vencidas — valor", '=SUMIFS(R_VALOR,R_STATUS,"A RECEBER",R_DATA,"<"&TODAY())', MOEDA0, None),
        ("Menor saldo de caixa projetado (exposição)", "=MIN(%s!$C$31:$%s$31)" % (FX, CL), MOEDA0, None),
        ("Soma das participações societárias", "=PartTotal", PCT, '=IF(ROUND(PartTotal,4)=1,"OK","ATENÇÃO")'),
        ("Sócios com aporte pendente (posição negativa)", '=SUMPRODUCT(--(APORTES!$M$6:$M$%d<-0.005))' % (5 + N_SOCIOS), "0", None),
    ]
    for i, (txt, fml, nf, st) in enumerate(alerts):
        r = 20 + i
        p.merge_range(r, 1, r, 10, txt, fm(border=1, border_color=LINE, indent=1))
        p.merge_range(r, 11, r, 13, "", f_au(nf, align="center", bold=True))
        p.write_formula(r, 11, fml, f_au(nf, align="center", bold=True))
        p.merge_range(r, 14, r, 16, "", f_au(align="center", bold=True))
        cell = "L%d" % (r + 1)
        if st is None:
            if "saldo" in txt:
                st = '=IF(%s<0,"ATENÇÃO","OK")' % cell
            else:
                st = '=IF(%s>0.005,"ATENÇÃO","OK")' % cell
        p.write_formula(r, 14, st, f_au(align="center", bold=True))
    al_area = "O21:Q%d" % (20 + len(alerts))
    p.conditional_format(al_area, {"type": "cell", "criteria": "==", "value": '"OK"', "format": fm(bg_color=GREEN_BG, font_color=GREEN_FG, bold=True)})
    p.conditional_format(al_area, {"type": "cell", "criteria": "==", "value": '"ATENÇÃO"', "format": fm(bg_color=RED_BG, font_color=RED_FG, bold=True)})

    # graficos
    cats = "=%s!$C$5:$%s$5" % (FX, CL)

    def ch_base(ch, title_):
        ch.set_title({"name": title_, "name_font": {"size": 11, "bold": True, "color": INK, "name": F_TIT}})
        ch.set_legend({"position": "bottom", "font": {"size": 8}})
        ch.set_chartarea({"border": {"color": CREAM}})
        return ch

    gr = 30
    p.merge_range("B%d:Q%d" % (gr, gr), "GRÁFICOS", f_sec)
    c1 = wb.add_chart({"type": "column"})
    c1.add_series({"name": "Entradas", "categories": cats, "values": "=%s!$C$15:$%s$15" % (FX, CL), "fill": {"color": SERIES[0]}, "gap": 60})
    c1.add_series({"name": "Saídas", "categories": cats, "values": "=%s!$C$24:$%s$24" % (FX, CL), "fill": {"color": SERIES[1]}})
    l1 = wb.add_chart({"type": "line"})
    l1.add_series({"name": "Saldo acumulado", "categories": cats, "values": "=CALC!$C$31:$%s$31" % CALC_CL,
                   "line": {"color": WARM_D, "width": 2.5}, "marker": {"type": "circle", "size": 4, "fill": {"color": WARM_D}, "border": {"color": WARM_D}}})
    c1.combine(l1)
    ch_base(c1, "Entradas × Saídas × Saldo acumulado")
    c1.set_x_axis({"num_format": MES, "num_font": {"size": 8, "rotation": -45}, "date_axis": False})
    c1.set_y_axis({"num_format": '"R$" #,##0,"k"', "num_font": {"size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#ECE7E0"}}})
    p.insert_chart("B%d" % (gr + 1), c1, {"x_scale": 1.27, "y_scale": 1.15})

    c2 = wb.add_chart({"type": "line"})
    c2.add_series({"name": "Planejado acumulado", "categories": cats, "values": "=%s!$C$39:$%s$39" % (FX, CL),
                   "line": {"color": WARM, "width": 2.25, "dash_type": "dash"}})
    c2.add_series({"name": "Realizado acumulado", "categories": cats, "values": "=CALC!$C$32:$%s$32" % CALC_CL,
                   "line": {"color": SERIES[0], "width": 2.75}})
    ch_base(c2, "Curva S — desembolso planejado × realizado")
    c2.set_x_axis({"num_format": MES, "num_font": {"size": 8, "rotation": -45}})
    c2.set_y_axis({"num_format": '"R$" #,##0,"k"', "num_font": {"size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#ECE7E0"}}})
    p.insert_chart("J%d" % (gr + 1), c2, {"x_scale": 1.27, "y_scale": 1.15})

    gr2 = gr + 19
    c3 = wb.add_chart({"type": "bar"})
    c3.add_series({"name": "Orçado", "categories": "=CALC!$O$2:$O$%d" % (1 + N_ETAPA), "values": "=CALC!$P$2:$P$%d" % (1 + N_ETAPA), "fill": {"color": WARM}, "gap": 50})
    c3.add_series({"name": "Realizado", "categories": "=CALC!$O$2:$O$%d" % (1 + N_ETAPA), "values": "=CALC!$Q$2:$Q$%d" % (1 + N_ETAPA), "fill": {"color": SERIES[0]}})
    ch_base(c3, "Orçado × Realizado por etapa")
    c3.set_y_axis({"reverse": True, "num_font": {"size": 7}})
    c3.set_x_axis({"num_format": '"R$" #,##0,"k"', "num_font": {"size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#ECE7E0"}}})
    p.insert_chart("B%d" % gr2, c3, {"x_scale": 1.27, "y_scale": 1.6})

    c4 = wb.add_chart({"type": "bar"})
    c4.add_series({"name": "% físico", "categories": "=CALC!$O$2:$O$%d" % (1 + N_ETAPA), "values": "=CALC!$S$2:$S$%d" % (1 + N_ETAPA), "fill": {"color": SERIES[2]}, "gap": 50})
    c4.add_series({"name": "% financeiro", "categories": "=CALC!$O$2:$O$%d" % (1 + N_ETAPA), "values": "=CALC!$R$2:$R$%d" % (1 + N_ETAPA), "fill": {"color": SERIES[1]}})
    ch_base(c4, "Avanço físico × financeiro por etapa")
    c4.set_y_axis({"reverse": True, "num_font": {"size": 7}})
    c4.set_x_axis({"num_format": "0%", "min": 0, "num_font": {"size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#ECE7E0"}}})
    p.insert_chart("J%d" % gr2, c4, {"x_scale": 1.27, "y_scale": 1.6})

    gr3 = gr2 + 25
    c5 = wb.add_chart({"type": "bar"})
    c5.add_series({"name": "Pago", "categories": "=CALC!$A$%d:$A$%d" % CALC_CAT, "values": "=CALC!$B$%d:$B$%d" % CALC_CAT,
                   "fill": {"color": SERIES[0]}, "gap": 40,
                   "data_labels": {"value": True, "num_format": '"R$" #,##0,"k";;', "font": {"size": 7}}})
    ch_base(c5, "Custos pagos por categoria")
    c5.set_legend({"none": True})
    c5.set_y_axis({"reverse": True, "num_font": {"size": 7}})
    c5.set_x_axis({"num_format": '"R$" #,##0,"k"', "num_font": {"size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#ECE7E0"}}})
    p.insert_chart("B%d" % gr3, c5, {"x_scale": 1.27, "y_scale": 1.3})

    c6 = wb.add_chart({"type": "column"})
    c6.add_series({"name": "Aportado", "categories": "=CALC!$D$2:$D$%d" % (1 + N_SOCIOS), "values": "=CALC!$E$2:$E$%d" % (1 + N_SOCIOS), "fill": {"color": SERIES[0]}, "gap": 80})
    c6.add_series({"name": "Cota líquida devida", "categories": "=CALC!$D$2:$D$%d" % (1 + N_SOCIOS), "values": "=CALC!$F$2:$F$%d" % (1 + N_SOCIOS), "fill": {"color": SERIES[1]}})
    ch_base(c6, "Aportes × cota de cada sócio")
    c6.set_y_axis({"num_format": '"R$" #,##0,"k"', "num_font": {"size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#ECE7E0"}}})
    c6.set_x_axis({"num_font": {"size": 8}})
    p.insert_chart("J%d" % gr3, c6, {"x_scale": 1.27, "y_scale": 1.3})

    gr4 = gr3 + 21
    c7 = wb.add_chart({"type": "pie"})
    c7.add_series({"name": "Oficial × Especial", "categories": "=CALC!$H$2:$H$4", "values": "=CALC!$I$2:$I$4",
                   "data_labels": {"percentage": True, "font": {"size": 9, "bold": True, "color": "white"}},
                   "points": [{"fill": {"color": SERIES[0]}}, {"fill": {"color": SERIES[1]}}, {"fill": {"color": WINE}}]})
    ch_base(c7, "Custo direto: caixa oficial × especial")
    p.insert_chart("B%d" % gr4, c7, {"x_scale": 1.27, "y_scale": 1.1})

    c8 = wb.add_chart({"type": "column"})
    c8.add_series({"name": "Saldo do mês", "categories": cats, "values": "=%s!$C$26:$%s$26" % (FX, CL),
                   "fill": {"color": SERIES[2]}, "invert_if_negative": True, "invert_if_negative_color": WINE, "gap": 40})
    ch_base(c8, "Saldo de caixa de cada mês")
    c8.set_legend({"none": True})
    c8.set_x_axis({"num_format": MES, "num_font": {"size": 8, "rotation": -45}})
    c8.set_y_axis({"num_format": '"R$" #,##0,"k"', "num_font": {"size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#ECE7E0"}}})
    p.insert_chart("J%d" % gr4, c8, {"x_scale": 1.27, "y_scale": 1.1})
    p.set_landscape()
    p.fit_to_pages(1, 0)
    p.activate()
    p.set_first_sheet()

    # =========================================================================
    # PAINEL COMERCIAL
    # =========================================================================
    pc = ws["PAINEL COMERCIAL"]
    pc.set_tab_color(BRAND)
    pc.set_column("A:A", 2)
    pc.set_column("B:Q", 10.5)
    pc.set_row(1, 34)
    pc.set_row(0, 22)
    pc.merge_range("B1:Q1", "HLM CONSTRUÇÕES  ·  COMERCIAL E VIABILIDADE", fm(font_name=F_MONO, font_size=8, bold=True, font_color=BRAND, valign="bottom"))
    pc.merge_range("B2:Q2", "", fm(bold=True, font_size=20, font_color="white", bg_color=INK, indent=1, font_name=F_TIT, bottom=5, bottom_color=BRAND))
    pc.write_formula("B2", '="PAINEL COMERCIAL  —  "&UPPER(NomeEmp)', fm(bold=True, font_size=20, font_color="white", bg_color=INK, indent=1, font_name=F_TIT, bottom=5, bottom_color=BRAND))
    pc.merge_range("B3:Q3", "", f_sub)
    pc.write_formula("B3", '=NumUnid&" unidades  |  "&ROUND(AreaPrivativa,0)&" m² privativos  |  Preço alvo: R$ "&ROUND(PrecoM2,0)&" por m²"', f_sub)
    pc.merge_range("B5:Q5", "VENDAS", f_sec)
    tile(pc, 5, 1, "VGV DE TABELA", "=vb_vgv_tab", MOEDA0, "Em estoque", "=vb_vgv_est", MOEDA0, INK)
    tile(pc, 5, 5, "VGV CONTRATADO", "=vb_vgv_vend", MOEDA0, "% do VGV", "=vb_pct_vend", PCT, WARM_D)
    tile(pc, 5, 9, "TICKET MÉDIO DE VENDA", '=IFERROR(vb_vgv_vend/COUNTIF(V_STATUS,"VENDIDA"),0)', MOEDA0, "Desconto médio s/ tabela",
         '=IFERROR(1-vb_vgv_vend/SUMIFS(V_TAB,V_STATUS,"VENDIDA"),0)', PCT, "#3A3633")
    tile(pc, 5, 13, "PREÇO MÉDIO VENDIDO (m²)", "=vb_pm2_vend", MOEDA0, "Tabela (m²)", "=vb_pm2_tab", MOEDA0, BRAND_D)
    pc.merge_range("B10:Q10", "RECEBÍVEIS E RESULTADO", f_sec)
    tile(pc, 10, 1, "RECEBIDO DE CLIENTES", '=SUMIFS(R_VALOR,R_STATUS,"RECEBIDO")', MOEDA0, "% dos contratos", '=IFERROR(SUMIFS(R_VALOR,R_STATUS,"RECEBIDO")/vb_vgv_vend,0)', PCT, BRAND_D)
    tile(pc, 10, 5, "CARTEIRA A RECEBER", '=SUMIFS(V_ARECEBER,V_STATUS,"VENDIDA")', MOEDA0, "Parcelas lançadas", '=SUMIFS(R_VALOR,R_STATUS,"A RECEBER")', MOEDA0, INK)
    tile(pc, 10, 9, "INADIMPLÊNCIA (vencido)", '=SUMIFS(R_VALOR,R_STATUS,"A RECEBER",R_DATA,"<"&TODAY())', MOEDA0, "Parcelas vencidas",
         '=COUNTIFS(R_STATUS,"A RECEBER",R_DATA,"<"&TODAY())', "0", WINE)
    tile(pc, 10, 13, "LUCRO PROJETADO", "=vb_lucro", MOEDA0, "ROI s/ custo", "=vb_roi", PCT, WARM_D)
    pc.merge_range("B15:Q15", "VIABILIDADE", f_sec)
    tile(pc, 15, 1, "RECEITA LÍQUIDA", "=vb_rec_liq", MOEDA0, "Comissões + impostos", "=vb_comis+vb_imp", MOEDA0, INK)
    tile(pc, 15, 5, "CUSTO TOTAL PROJETADO", "=vb_custo_tot", MOEDA0, "Custo / m² privativo", "=vb_cm2p", MOEDA0, BRAND)
    tile(pc, 15, 9, "PONTO DE EQUILÍBRIO (VGV)", "=vb_be_vgv", MOEDA0, "Preço mín. / m²", "=vb_be_m2", MOEDA0, "#3A3633")
    tile(pc, 15, 13, "EXPOSIÇÃO MÁXIMA DE CAIXA", "=vb_expo", MOEDA0, "Avanço físico", "=vb_pct_fis", PCT, BRAND_D)

    gq = 20
    pc.merge_range("B%d:Q%d" % (gq, gq), "GRÁFICOS", f_sec)
    q1 = wb.add_chart({"type": "doughnut"})
    q1.add_series({"name": "Unidades por status", "categories": "=CALC!$K$2:$K$5", "values": "=CALC!$L$2:$L$5",
                   "data_labels": {"value": True, "font": {"bold": True, "color": "white"}},
                   "points": [{"fill": {"color": WARM}}, {"fill": {"color": WARM_D}}, {"fill": {"color": BRAND}}, {"fill": {"color": INK}}]})
    q1.set_hole_size(50)
    ch_base(q1, "Unidades por status")
    pc.insert_chart("B%d" % (gq + 1), q1, {"x_scale": 1.27, "y_scale": 1.15})
    q2 = wb.add_chart({"type": "column"})
    q2.add_series({"name": "Recebido", "categories": cats, "values": "=%s!$C$14:$%s$14" % (FX, CL), "fill": {"color": BRAND}, "gap": 50})
    q2.add_series({"name": "A receber (previsto)", "categories": cats, "values": "=%s!$C$30:$%s$30" % (FX, CL), "fill": {"color": CREAM}})
    ch_base(q2, "Recebimentos de clientes por mês")
    q2.set_x_axis({"num_format": MES, "num_font": {"size": 8, "rotation": -45}})
    q2.set_y_axis({"num_format": '"R$" #,##0,"k"', "num_font": {"size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#ECE7E0"}}})
    pc.insert_chart("J%d" % (gq + 1), q2, {"x_scale": 1.27, "y_scale": 1.15})
    q3 = wb.add_chart({"type": "column"})
    cr0 = VI["cen_r0"]
    q3.add_series({"name": "Lucro", "categories": "=VIABILIDADE!$F$%d:$F$%d" % (cr0 + 2, cr0 + 4),
                   "values": "=VIABILIDADE!$I$%d:$I$%d" % (cr0 + 2, cr0 + 4),
                   "points": [{"fill": {"color": WINE}}, {"fill": {"color": INK}}, {"fill": {"color": BRAND}}],
                   "data_labels": {"value": True, "num_format": '"R$" #,##0,"k"', "font": {"size": 8}}, "gap": 80})
    ch_base(q3, "Lucro projetado por cenário de preço")
    q3.set_legend({"none": True})
    q3.set_y_axis({"num_format": '"R$" #,##0,"k"', "num_font": {"size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#ECE7E0"}}})
    pc.insert_chart("B%d" % (gq + 19), q3, {"x_scale": 1.27, "y_scale": 1.15})
    q4 = wb.add_chart({"type": "pie"})
    q4.add_series({"name": "VGV", "categories": "=CALC!$K$2:$K$4", "values": "=CALC!$M$2:$M$4",
                   "data_labels": {"percentage": True, "font": {"bold": True, "color": "white"}},
                   "points": [{"fill": {"color": WARM}}, {"fill": {"color": WARM_D}}, {"fill": {"color": BRAND}}]})
    ch_base(q4, "VGV: estoque × reservado × vendido")
    pc.insert_chart("J%d" % (gq + 19), q4, {"x_scale": 1.27, "y_scale": 1.15})

    # espelho das unidades
    tq = gq + 38
    pc.merge_range("B%d:Q%d" % (tq, tq), "ESPELHO DE VENDAS", f_sec)
    hdrs = [("UNIDADE", 1), ("STATUS", 2), ("CLIENTE", 4), ("ÁREA m²", 1), ("TABELA", 2), ("VENDA", 2), ("RECEBIDO", 2), ("% REC.", 2)]
    col = 1
    for h, span in hdrs:
        if span > 1:
            pc.merge_range(tq, col, tq, col + span - 1, h, f_hdr)
        else:
            pc.write(tq, col, h, f_hdr)
        col += span
    srcs = [("A", None, 1), ("F", None, 2), ("G", None, 4), ("C", "#,##0.0", 1), ("E", MOEDA0, 2), ("K", MOEDA0, 2), ("O", MOEDA0, 2), ("Q", PCT, 2)]
    for k in range(V_FIM - V_INI + 1):
        r = tq + 1 + k
        vr = V_INI + k
        col = 1
        for src, nf, span in srcs:
            fmt = fm(border=1, border_color=LINE, font_size=9, num_format=nf) if nf else fm(border=1, border_color=LINE, font_size=9)
            fml = '=IF(VENDAS!$A${0}="","",VENDAS!{1}{0}&"")'.format(vr, src) if nf is None else '=IF(VENDAS!$A${0}="","",VENDAS!{1}{0})'.format(vr, src)
            if span > 1:
                pc.merge_range(r, col, r, col + span - 1, "", fmt)
            pc.write_formula(r, col, fml, fmt)
            col += span
    pc.conditional_format("C%d:D%d" % (tq + 2, tq + 1 + V_FIM - V_INI + 1), {"type": "cell", "criteria": "==", "value": '"VENDIDA"', "format": fm(bg_color=GREEN_BG, font_color=GREEN_FG, bold=True)})
    pc.conditional_format("C%d:D%d" % (tq + 2, tq + 1 + V_FIM - V_INI + 1), {"type": "cell", "criteria": "==", "value": '"RESERVADA"', "format": fm(bg_color=AMB_BG, font_color=AMB_FG, bold=True)})
    pc.conditional_format("C%d:D%d" % (tq + 2, tq + 1 + V_FIM - V_INI + 1), {"type": "cell", "criteria": "==", "value": '"PERMUTA"', "format": fm(bg_color="#ECE7E0", font_color="#3A3633", bold=True)})
    pc.set_landscape()
    pc.fit_to_pages(1, 0)

    # =========================================================================
    # GUIA
    # =========================================================================
    gu = ws["GUIA"]
    gu.set_tab_color(WARM_D)
    gu.set_column("A:A", 2)
    gu.set_column("B:B", 30)
    gu.set_column("C:C", 110)
    gu.set_row(1, 30)
    gu.merge_range("B2:C2", "", f_title)
    gu.write_formula("B2", nome_titulo("GUIA DE USO"), f_title)
    gu.merge_range("B3:C3", "Como a planilha funciona, onde preencher e o que foi corrigido em relação ao modelo HLM ONE.", f_sub)
    f_gk = fm(bold=True, border=1, border_color=LINE, indent=1, valign="top", text_wrap=True)
    f_gv = fm(border=1, border_color=LINE, indent=1, valign="top", text_wrap=True)
    secs = [
        ("IDENTIDADE VISUAL HLM", [
            ("Cores", "Preto HLM #040203 · Laranja HLM #D45508 (escuro #A83F03) · Creme #DBD4CA · Cinza quente #AEA79F · Vinho #951D25 (alertas)."),
            ("Tipografia", "Títulos: Space Grotesk · Rótulos: Space Mono · Dados: Arial. Space Grotesk e Space Mono são gratuitas (Google Fonts); "
                           "no Google Planilhas já estão disponíveis. No Excel, instale-as para ver a tipografia exata (sem elas, o Excel usa uma fonte padrão)."),
        ]),
        ("LEGENDA", [
            ("Célula LARANJA-CLARA", "Você preenche."),
            ("Célula CINZA", "Calculada automaticamente — não digite por cima."),
            ("Coluna VERIFICAÇÃO", "OK em verde = lançamento válido. Vermelho = o que está errado naquela linha."),
        ]),
        ("ORDEM DE PREENCHIMENTO", [
            ("1. CONFIG", "Nome do empreendimento, data de início, prazo, taxa de adm, sócios e % de participação, categorias, etapas, equipe."),
            ("2. ORÇAMENTO", "Valor orçado de cada etapa. Atualize o % de AVANÇO FÍSICO mensalmente."),
            ("3. VENDAS", "Todas as unidades (área, tipologia). Ao vender: STATUS=VENDIDA, cliente, data e valor."),
            ("4. Dia a dia", "DESPESAS (todo gasto), FOLHA PAG (quinzenas), APORTES (depósitos dos sócios), RECEBIMENTOS (parcelas dos clientes)."),
            ("5. Fechamento do mês", "Confira a AUDITORIA (tudo OK) → ACERTO MENSAL → escolha o mês → BAIXAR ACERTO (PDF) → envie aos sócios."),
        ]),
        ("COMO OS DADOS SE CONECTAM", [
            ("DESPESAS", "→ FLUXO DE CAIXA (oficial/especial, taxa de adm) → ACERTO (por categoria e detalhamento) → ORÇAMENTO (por etapa) → PAINEL (composição) → AUDITORIA."),
            ("FOLHA PAG", "→ FLUXO (folha oficial/especial + taxa de adm) → ACERTO (seção 5) → ORÇAMENTO (por etapa)."),
            ("APORTES", "→ FLUXO (entradas por sócio) → ACERTO (aportes do mês e posição acumulada) → APORTES (resumo) → VIABILIDADE (ROI por sócio)."),
            ("VENDAS", "→ VIABILIDADE (VGV, comissões, impostos, lucro) → PAINEL COMERCIAL → RECEBIMENTOS (cliente/unidade)."),
            ("RECEBIMENTOS", "→ FLUXO (entradas e previsão) → ACERTO (cota das receitas de cada sócio) → VENDAS (recebido por unidade)."),
            ("CONFIG", "→ todas as listas suspensas, % de rateio, taxa de adm, preço/m², títulos de todas as abas."),
            ("ORÇAMENTO", "→ curva S planejada (FLUXO l.38) → custo projetado (VIABILIDADE) → alertas do PAINEL."),
        ]),
        ("REGRAS DE CÁLCULO", [
            ("Mês", "Sempre calculado a partir da DATA do lançamento (antes era digitado à parte e podia divergir)."),
            ("Taxa de administração", "TaxaAdm × (despesas pagas de categorias com INCIDE ADM = SIM + folha, se marcado). Comissões, marketing e impostos de venda não pagam taxa."),
            ("Rateio", "Pela % de participação de cada sócio em CONFIG (antes era ÷4 fixo)."),
            ("Posição do sócio", "Aportes acumulados + cota das receitas − cota dos custos c/ adm. Negativo = deve aportar; positivo = crédito."),
            ("A PAGAR / A RECEBER", "Não entram no saldo realizado; aparecem na previsão (FLUXO l.29–31) pelo vencimento."),
            ("Custo projetado (EAC)", "Para etapas com avanço físico ≥ 10%: realizado ÷ % físico (nunca menor que o já gasto + comprometido). Abaixo disso: o orçado."),
        ]),
        ("BAIXAR O ACERTO", [
            ("Excel (arquivo .xlsm)", "Habilite as macros. Na aba ACERTO MENSAL, escolha o mês e clique em BAIXAR ACERTO (PDF). As linhas vazias do detalhamento são ocultadas no PDF. "
                                      "Se o Windows bloquear macros de arquivo baixado: botão direito no arquivo › Propriedades › marcar 'Desbloquear'."),
            ("Google Planilhas / .xlsx", "Abra a aba ACERTO MENSAL › Arquivo › Fazer download › PDF › Exportar: 'Página atual'. A área de impressão já está configurada (A4 retrato, 1 página de largura)."),
        ]),
        ("CORREÇÕES EM RELAÇÃO AO HLM ONE", [
            ("1. Mês digitado à parte", "O MÊS/PER. era escolhido manualmente, independente da DATA → lançamento com data de outubro podia cair em setembro. Agora é calculado."),
            ("2. Calendário fixo", "Lista de meses fixa Set/25–Set/26 (13 meses). Agora o fluxo tem %d meses a partir da data de início em CONFIG." % N_MESES),
            ("3. Nomes de categoria divergentes", "Na versão de teste a lista tinha 'Projetos' e 'IPTU e Taxas', mas o ACERTO/ORÇAMENTO somavam 'Projetos, IPTU e Taxas' → sempre zero. Agora todas as abas leem a mesma lista de CONFIG."),
            ("4. Aporte/saldo por sócio ÷4", "O DASHBOARD dividia total aportado e saldo por 4, ignorando quanto cada sócio realmente aportou. Agora é por sócio e por % de participação."),
            ("5. Referências quebradas", "DASHBOARD (teste) apontava para FLUXO!P20 e P24, células vazias → 'gasto por sócio' e 'saldo por sócio' sempre zero. Corrigido."),
            ("6. Gráfico com 26 séries", "Gráfico de entradas × saídas tinha uma série por célula (26 séries soltas). Agora são séries corretas por mês."),
            ("7. Lançamentos sem TIPO sumiam", "Despesa sem OFICIAL/ESPECIAL não entrava no fluxo nem no acerto. Agora aparece numa linha 'corrigir' e na AUDITORIA."),
            ("8. 'Outros' por diferença", "'Outros Gastos Oficiais' era calculado por diferença e escondia categorias erradas. Agora a diferença vira 'sem categoria (corrigir)'."),
            ("9. Ajustes manuais fixos", "FLUXO (teste) tinha 'ajustes manuais' de R$ 100.000 digitados nos meses, inflando as saídas. Removido: todo gasto passa por DESPESAS."),
            ("10. Taxa 15% e ÷4 fixos nas fórmulas", "Agora são parâmetros únicos em CONFIG (TaxaAdm e % de participação)."),
            ("11. Folha em locais diferentes", "A folha era 'caixa especial' no ACERTO mas linha separada no FLUXO. Agora cada pagamento tem CAIXA (oficial/especial) e os dois relatórios batem."),
            ("12. Saldo anterior manual", "O ACERTO pedia o saldo anterior digitado. Agora vem do saldo acumulado do FLUXO, com conferência ✔."),
            ("13. Lucro por sócio ÷4", "COMERCIAL dividia lucro e aporte por 4. Agora usa a % de cada sócio e mostra ROI sobre o aporte."),
            ("14. Custo comercial misturado", "Comissões/impostos de venda eram somados ao custo da obra e pagavam taxa de adm. Agora ficam fora do custo da obra e da taxa."),
        ]),
    ]
    r = 4
    for sec, items in secs:
        gu.merge_range(r, 1, r, 2, sec, f_sec)
        r += 1
        for kx, vx in items:
            gu.write(r, 1, kx, f_gk)
            gu.write(r, 2, vx, f_gv)
            gu.set_row(r, 30 if len(vx) > 110 else 18)
            r += 1
        r += 1

    ws["_guia"] = (gu, r, f_sec, f_gk, f_gv)
    if dados:
        dados(ws, wb, fm)   # preenchimento opcional com dados reais (módulo local, fora do Git)
    wb.close()


def main():
    os.makedirs(SAIDA, exist_ok=True)
    base = os.path.join(SAIDA, "HLM_Gestao_Novo_Empreendimento")
    build(base + ".xlsm", macro=True)
    build(base + ".xlsx")
    # Planilha do HLM ONE com dados reais: só é gerada se o módulo local scripts/hlm_one/preencher.py
    # e o CSV dados/hlm_one_lancamentos.csv existirem (ambos ficam fora do Git por conterem dados financeiros reais).
    try:
        from hlm_one.preencher import preencher_hlm_one
    except ImportError:
        preencher_hlm_one = None
    if preencher_hlm_one and os.path.exists(DADOS_HLM_ONE):
        build(os.path.join(SAIDA, "HLM_ONE_Gestao.xlsm"), macro=True,
              dados=lambda ws, wb, fm: preencher_hlm_one(ws, wb, fm, DADOS_HLM_ONE))
    print("ok:", os.listdir(SAIDA))


if __name__ == "__main__":
    main()
