"""Gera a planilha de pratica do tutorial (dados ficticios).

Uso:  python ferramentas/gerar_planilha.py
Saida: arquivos/planilha-pratica-recrutamento.xlsx

As formulas sao gravadas com nomes em ingles (padrao do arquivo .xlsx);
o Excel em portugues mostra os nomes traduzidos automaticamente
(COUNTIFS -> CONT.SES, IFERROR -> SEERRO etc.).
"""
from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

random.seed(42)

AZUL = "1D4ED8"
AZUL_ESCURO = "1E3A8A"
AZUL_CLARO = "DBEAFE"
BRANCO = "FFFFFF"

ETAPAS = ["Inscrito", "Triagem", "Entrevista RH", "Teste técnico",
          "Entrevista gestor", "Proposta", "Contratado"]
STATUS = ["Em andamento", "Reprovado", "Desistiu", "Contratado", "Banco de talentos"]
FONTES = ["LinkedIn", "Indicação", "Site de carreiras", "Indeed", "Hunting",
          "Banco de talentos", "Universidades"]
RECRUTADORES = ["Ana Lima", "Bruno Costa", "Carla Dias"]
CIDADES = ["São Paulo", "Campinas", "Rio de Janeiro", "Belo Horizonte", "Curitiba", "Remoto"]

VAGAS = [
    # codigo, cargo, area, gestor, abertura, sla, status, fechamento
    ("VG-001", "Analista de Dados Pl", "Tecnologia", "Marcos Prado", date(2026, 8, 3), 30, "Fechada", date(2026, 9, 18)),
    ("VG-002", "Assistente Administrativo", "Administrativo", "Paula Reis", date(2026, 8, 17), 20, "Fechada", date(2026, 9, 9)),
    ("VG-003", "Desenvolvedor Back-end Sr", "Tecnologia", "Marcos Prado", date(2026, 9, 1), 45, "Aberta", None),
    ("VG-004", "Analista de RH Jr", "Pessoas", "Lúcia Mendes", date(2026, 9, 14), 25, "Aberta", None),
    ("VG-005", "Executivo de Vendas", "Comercial", "Rafael Gomes", date(2026, 9, 21), 30, "Aberta", None),
    ("VG-006", "Coordenador Financeiro", "Financeiro", "Sônia Alves", date(2026, 10, 5), 40, "Aberta", None),
    ("VG-007", "Estagiário de Marketing", "Marketing", "Tiago Nunes", date(2026, 10, 19), 20, "Congelada", None),
]

NOMES = ["Mariana", "João", "Fernanda", "Lucas", "Beatriz", "Gabriel", "Juliana", "Pedro",
         "Camila", "Rafael", "Larissa", "Thiago", "Aline", "Diego", "Patrícia", "Felipe",
         "Renata", "Gustavo", "Vanessa", "André", "Letícia", "Rodrigo", "Bianca", "Eduardo",
         "Natália", "Marcelo", "Priscila", "Vinícius", "Tatiane", "Leonardo"]
SOBRENOMES = ["Silva", "Santos", "Oliveira", "Souza", "Pereira", "Almeida", "Ferreira",
              "Rodrigues", "Carvalho", "Gomes", "Martins", "Araújo", "Ribeiro", "Barbosa",
              "Rocha", "Moreira", "Cardoso", "Teixeira", "Correia", "Melo"]

fino = Side(style="thin", color="CBD5E1")
borda = Border(left=fino, right=fino, top=fino, bottom=fino)


def sem_acento(txt: str) -> str:
    tabela = str.maketrans("áàâãéêíóôõúçÁÀÂÃÉÊÍÓÔÕÚÇ", "aaaaeeiooouc" "AAAAEEIOOOUC")
    return txt.translate(tabela)


def titulo(ws, texto: str, sub: str | None = None) -> None:
    ws["A1"] = texto
    ws["A1"].font = Font(bold=True, size=16, color=AZUL_ESCURO)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, size=10, color="475569")


def larguras(ws, valores: dict[str, float]) -> None:
    for col, larg in valores.items():
        ws.column_dimensions[col].width = larg


def gerar_candidatos(qtd: int = 60) -> list[dict]:
    hoje_ref = date(2026, 11, 20)
    usados: set[str] = set()
    linhas = []
    for i in range(1, qtd + 1):
        while True:
            nome = f"{random.choice(NOMES)} {random.choice(SOBRENOMES)}"
            if nome not in usados:
                usados.add(nome)
                break
        vaga = random.choice(VAGAS)
        abertura = vaga[4]
        inscricao = abertura + timedelta(days=random.randint(0, 20))
        if inscricao > hoje_ref:
            inscricao = hoje_ref - timedelta(days=random.randint(1, 5))

        status = random.choices(STATUS, weights=[40, 30, 10, 0, 10])[0]
        n_etapa = random.choices(range(1, 7), weights=[15, 25, 22, 15, 13, 10])[0]
        fechada = vaga[6] == "Fechada"
        if fechada and status == "Em andamento":
            status = random.choice(["Reprovado", "Banco de talentos"])

        ultima = inscricao + timedelta(days=random.randint(1, 4) * n_etapa)
        if ultima > hoje_ref:
            ultima = hoje_ref - timedelta(days=random.randint(0, 3))
        contratacao = None
        linhas.append({
            "id": f"CAND-{i:03d}", "nome": nome,
            "email": f"{sem_acento(nome).lower().replace(' ', '.')}@email.com",
            "tel": f"(11) 9{random.randint(1000, 9999)}-{random.randint(1000, 9999)}",
            "cidade": random.choice(CIDADES), "vaga": vaga[0],
            "fonte": random.choices(FONTES, weights=[30, 18, 15, 12, 8, 10, 7])[0],
            "recrutador": random.choice(RECRUTADORES), "etapa": ETAPAS[n_etapa - 1],
            "status": status, "inscricao": inscricao, "ultima": ultima,
            "contratacao": contratacao,
            "pretensao": random.choice(range(2500, 16001, 250)),
            "nota": round(random.uniform(2.5, 5), 1) if n_etapa >= 3 else None,
        })

    # um contratado por vaga fechada
    for vaga in VAGAS:
        if vaga[6] != "Fechada":
            continue
        cand = next(c for c in linhas if c["vaga"] == vaga[0] and c["status"] != "Contratado")
        cand.update(etapa="Contratado", status="Contratado", contratacao=vaga[7],
                    ultima=vaga[7], nota=round(random.uniform(4.2, 5), 1),
                    inscricao=min(cand["inscricao"], vaga[7] - timedelta(days=15)))
    return linhas


def aba_listas(wb: Workbook) -> None:
    ws = wb.create_sheet("Listas")
    titulo(ws, "Listas de apoio", "Usadas nas listas suspensas (Dados > Validação de Dados). Edite aqui para mudar as opções.")
    colunas = {"A": ("Etapas", ETAPAS), "C": ("Status", STATUS), "E": ("Fontes", FONTES),
               "G": ("Recrutadores", RECRUTADORES), "I": ("Feriados", [
                   date(2026, 9, 7), date(2026, 10, 12), date(2026, 11, 2),
                   date(2026, 11, 15), date(2026, 11, 20), date(2026, 12, 25)])}
    for col, (cab, valores) in colunas.items():
        c = ws[f"{col}4"]
        c.value = cab
        c.font = Font(bold=True, color=BRANCO)
        c.fill = PatternFill("solid", fgColor=AZUL)
        for j, v in enumerate(valores, start=5):
            cel = ws[f"{col}{j}"]
            cel.value = v
            if isinstance(v, date):
                cel.number_format = "dd/mm/yyyy"
        ws.column_dimensions[col].width = 22
    ws["K4"] = "Ordem da etapa"
    ws["K4"].font = Font(bold=True, color=BRANCO)
    ws["K4"].fill = PatternFill("solid", fgColor=AZUL)
    for j in range(len(ETAPAS)):
        ws[f"K{5 + j}"] = j + 1
    ws.column_dimensions["K"].width = 16


def aba_vagas(wb: Workbook) -> None:
    ws = wb.create_sheet("Vagas")
    cab = ["Código", "Cargo", "Área", "Gestor", "Data Abertura", "SLA (dias úteis)",
           "Status Vaga", "Data Fechamento", "Dias úteis em aberto", "SLA estourado?"]
    ws.append(cab)
    for i, v in enumerate(VAGAS, start=2):
        ws.append(list(v))
        ws[f"I{i}"] = (f'=IF(E{i}="","",NETWORKDAYS(E{i},IF(H{i}="",TODAY(),H{i}),'
                       f'Listas!$I$5:$I$10))')
        ws[f"J{i}"] = f'=IF(I{i}="","",IF(I{i}>F{i},"Sim","Não"))'
        for col in "EH":
            ws[f"{col}{i}"].number_format = "dd/mm/yyyy"
    ultima = len(VAGAS) + 1
    tab = Table(displayName="tbl_Vagas", ref=f"A1:J{ultima}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    larguras(ws, {"A": 10, "B": 28, "C": 16, "D": 16, "E": 14, "F": 16, "G": 13,
                  "H": 16, "I": 20, "J": 15})
    ws.conditional_formatting.add(
        f"J2:J{ultima}", CellIsRule(operator="equal", formula=['"Sim"'],
                                    fill=PatternFill("solid", fgColor="FECACA"),
                                    font=Font(bold=True, color="991B1B")))
    dv = DataValidation(type="list", formula1='"Aberta,Fechada,Congelada,Cancelada"', allow_blank=False)
    dv.error = "Escolha um status da lista."
    dv.errorTitle = "Status inválido"
    ws.add_data_validation(dv)
    dv.add(f"G2:G{ultima + 50}")
    ws.freeze_panes = "B2"


def aba_candidatos(wb: Workbook, cands: list[dict]) -> int:
    ws = wb.create_sheet("Candidatos", 0)
    cab = ["ID", "Nome", "E-mail", "Telefone", "Cidade", "Cód. Vaga", "Cargo", "Fonte",
           "Recrutador", "Etapa", "Nº Etapa", "Status", "Data Inscrição",
           "Última Movimentação", "Data Contratação", "Pretensão Salarial",
           "Nota Entrevista", "Dias no processo", "Dias parado"]
    ws.append(cab)
    for i, c in enumerate(cands, start=2):
        ws.append([c["id"], c["nome"], c["email"], c["tel"], c["cidade"], c["vaga"], None,
                   c["fonte"], c["recrutador"], c["etapa"], None, c["status"], c["inscricao"],
                   c["ultima"], c["contratacao"], c["pretensao"], c["nota"], None, None])
        ws[f"G{i}"] = f'=IFERROR(INDEX(Vagas!$B$2:$B$100,MATCH(F{i},Vagas!$A$2:$A$100,0)),"Vaga não encontrada")'
        ws[f"K{i}"] = f'=IFERROR(MATCH(J{i},Listas!$A$5:$A$11,0),"")'
        ws[f"R{i}"] = (f'=IF(M{i}="","",NETWORKDAYS(M{i},IF(O{i}="",TODAY(),O{i}),'
                       f'Listas!$I$5:$I$10))')
        ws[f"S{i}"] = f'=IF(L{i}<>"Em andamento","",NETWORKDAYS(N{i},TODAY(),Listas!$I$5:$I$10)-1)'
        for col in "MNO":
            ws[f"{col}{i}"].number_format = "dd/mm/yyyy"
        ws[f"P{i}"].number_format = 'R$ #,##0.00'
        ws[f"Q{i}"].number_format = "0.0"
    ultima = len(cands) + 1
    tab = Table(displayName="tbl_Candidatos", ref=f"A1:S{ultima}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    larguras(ws, {"A": 11, "B": 24, "C": 32, "D": 16, "E": 16, "F": 11, "G": 27, "H": 18,
                  "I": 14, "J": 18, "K": 10, "L": 18, "M": 15, "N": 20, "O": 17, "P": 18,
                  "Q": 16, "R": 17, "S": 13})
    ws.freeze_panes = "C2"

    limite = ultima + 200  # validacao ja preparada para novas linhas
    for col, origem, msg in [("H", "=Listas!$E$5:$E$11", "fonte"),
                             ("I", "=Listas!$G$5:$G$7", "recrutador"),
                             ("J", "=Listas!$A$5:$A$11", "etapa"),
                             ("L", "=Listas!$C$5:$C$9", "status")]:
        dv = DataValidation(type="list", formula1=origem, allow_blank=True)
        dv.errorTitle = "Valor fora da lista"
        dv.error = f"Escolha uma {msg} da lista. Para incluir opções, edite a aba Listas."
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{limite}")
    dv_nota = DataValidation(type="decimal", operator="between", formula1="0", formula2="5",
                             allow_blank=True)
    dv_nota.errorTitle = "Nota inválida"
    dv_nota.error = "A nota deve ficar entre 0 e 5."
    ws.add_data_validation(dv_nota)
    dv_nota.add(f"Q2:Q{limite}")
    dv_data = DataValidation(type="date", operator="greaterThan", formula1="DATE(2020,1,1)",
                             allow_blank=True)
    dv_data.errorTitle = "Data inválida"
    dv_data.error = "Digite uma data válida (dd/mm/aaaa)."
    ws.add_data_validation(dv_data)
    dv_data.add(f"M2:O{limite}")

    faixa = f"A2:S{ultima}"
    ws.conditional_formatting.add(faixa, FormulaRule(
        formula=['$L2="Contratado"'], fill=PatternFill("solid", fgColor="DCFCE7")))
    ws.conditional_formatting.add(faixa, FormulaRule(
        formula=['AND($L2="Em andamento",$S2>=5)'], fill=PatternFill("solid", fgColor="FEF3C7")))
    ws.conditional_formatting.add(f"Q2:Q{ultima}", DataBarRule(
        start_type="num", start_value=0, end_type="num", end_value=5, color="3B82F6"))
    return ultima


def aba_painel(wb: Workbook, ultima: int) -> None:
    ws = wb.create_sheet("Painel", 1)
    titulo(ws, "Painel de Recrutamento", "Tudo calculado com fórmulas a partir das abas Candidatos e Vagas. Dados fictícios.")
    C = lambda col: f"Candidatos!${col}$2:${col}${ultima}"  # noqa: E731

    ws["A4"] = "Indicador"
    ws["B4"] = "Valor"
    ws["C4"] = "Como ler"
    indicadores = [
        ("Total de candidatos", f"=COUNTA({C('A')})", "0", "Quantas pessoas entraram no funil."),
        ("Em andamento", f'=COUNTIFS({C("L")},"Em andamento")', "0", "Processos vivos agora."),
        ("Contratados", f'=COUNTIFS({C("L")},"Contratado")', "0", "Vagas preenchidas."),
        ("Parados há 5+ dias úteis", f'=COUNTIFS({C("L")},"Em andamento",{C("S")},">=5")', "0",
         "Candidatos esperando retorno: priorize."),
        ("Time to hire (dias úteis, média)", f'=IFERROR(AVERAGEIFS({C("R")},{C("L")},"Contratado"),"-")',
         "0.0", "Da inscrição à contratação."),
        ("Time to fill (dias úteis, média)", '=IFERROR(AVERAGEIFS(Vagas!$I$2:$I$100,Vagas!$G$2:$G$100,"Fechada"),"-")',
         "0.0", "Da abertura da vaga ao fechamento."),
        ("Vagas abertas", '=COUNTIFS(Vagas!$G$2:$G$100,"Aberta")', "0", "Demanda atual."),
        ("Vagas com SLA estourado", '=COUNTIFS(Vagas!$G$2:$G$100,"Aberta",Vagas!$J$2:$J$100,"Sim")', "0",
         "Converse com o gestor sobre prazo ou perfil."),
        ("Nota média de entrevista", f'=IFERROR(AVERAGE({C("Q")}),"-")', "0.0", "Escala de 0 a 5."),
        ("Pretensão salarial média", f'=IFERROR(AVERAGE({C("P")}),"-")', 'R$ #,##0', "Referência para negociação."),
    ]
    for i, (nome, f, fmt, ajuda) in enumerate(indicadores, start=5):
        ws[f"A{i}"] = nome
        ws[f"B{i}"] = f
        ws[f"B{i}"].number_format = fmt
        ws[f"B{i}"].font = Font(bold=True, size=12, color=AZUL_ESCURO)
        ws[f"C{i}"] = ajuda
        ws[f"C{i}"].font = Font(color="475569")

    # funil
    ini = 5 + len(indicadores) + 2
    ws[f"A{ini}"] = "Funil (chegaram pelo menos até a etapa)"
    ws[f"A{ini}"].font = Font(bold=True, size=13, color=AZUL_ESCURO)
    ws[f"A{ini + 1}"] = "Etapa"
    ws[f"B{ini + 1}"] = "Candidatos"
    ws[f"C{ini + 1}"] = "Conversão da etapa anterior"
    for j, etapa in enumerate(ETAPAS):
        lin = ini + 2 + j
        ws[f"A{lin}"] = etapa
        ws[f"B{lin}"] = f'=COUNTIFS({C("K")},">="&{j + 1})'
        if j == 0:
            ws[f"C{lin}"] = "-"
        else:
            ws[f"C{lin}"] = f'=IFERROR(B{lin}/B{lin - 1},0)'
            ws[f"C{lin}"].number_format = "0.0%"
    fim_funil = ini + 1 + len(ETAPAS)
    ws.conditional_formatting.add(f"B{ini + 2}:B{fim_funil}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color="2563EB"))

    # por fonte
    ini2 = fim_funil + 3
    ws[f"A{ini2}"] = "Desempenho por fonte"
    ws[f"A{ini2}"].font = Font(bold=True, size=13, color=AZUL_ESCURO)
    for k, txt in enumerate(["Fonte", "Candidatos", "Chegaram à entrevista RH", "Contratados", "% que avançou"]):
        ws.cell(row=ini2 + 1, column=k + 1, value=txt)
    for j, fonte in enumerate(FONTES):
        lin = ini2 + 2 + j
        ws[f"A{lin}"] = fonte
        ws[f"B{lin}"] = f'=COUNTIFS({C("H")},A{lin})'
        ws[f"C{lin}"] = f'=COUNTIFS({C("H")},A{lin},{C("K")},">=3")'
        ws[f"D{lin}"] = f'=COUNTIFS({C("H")},A{lin},{C("L")},"Contratado")'
        ws[f"E{lin}"] = f'=IFERROR(C{lin}/B{lin},0)'
        ws[f"E{lin}"].number_format = "0.0%"

    # por recrutador
    ini3 = ini2 + 2 + len(FONTES) + 2
    ws[f"A{ini3}"] = "Carteira por recrutador"
    ws[f"A{ini3}"].font = Font(bold=True, size=13, color=AZUL_ESCURO)
    for k, txt in enumerate(["Recrutador", "Em andamento", "Parados 5+ dias", "Contratados"]):
        ws.cell(row=ini3 + 1, column=k + 1, value=txt)
    for j, rec in enumerate(RECRUTADORES):
        lin = ini3 + 2 + j
        ws[f"A{lin}"] = rec
        ws[f"B{lin}"] = f'=COUNTIFS({C("I")},A{lin},{C("L")},"Em andamento")'
        ws[f"C{lin}"] = f'=COUNTIFS({C("I")},A{lin},{C("L")},"Em andamento",{C("S")},">=5")'
        ws[f"D{lin}"] = f'=COUNTIFS({C("I")},A{lin},{C("L")},"Contratado")'

    for lin in (4, ini + 1, ini2 + 1, ini3 + 1):
        for col in range(1, 6):
            cel = ws.cell(row=lin, column=col)
            if cel.value is not None:
                cel.font = Font(bold=True, color=BRANCO)
                cel.fill = PatternFill("solid", fgColor=AZUL)
    larguras(ws, {"A": 36, "B": 16, "C": 44, "D": 14, "E": 15})


def aba_brutos(wb: Workbook) -> None:
    ws = wb.create_sheet("Dados_Brutos")
    titulo(ws, "Exercício de limpeza", "Dados como chegam de formulário ou exportação. Preencha as colunas em azul com fórmulas (veja o módulo Limpeza do tutorial).")
    cab = ["Nome (bruto)", "E-mail (bruto)", "Telefone (bruto)", "Cidade/UF (bruto)",
           "Nome limpo", "E-mail limpo", "Só dígitos do telefone", "Cidade", "UF", "E-mail parece válido?"]
    for k, txt in enumerate(cab, start=1):
        cel = ws.cell(row=4, column=k, value=txt)
        cel.font = Font(bold=True, color=BRANCO)
        cel.fill = PatternFill("solid", fgColor=AZUL if k <= 4 else "3B82F6")
    brutos = [
        ("  maria   eduarda SANTOS ", " Maria.Santos@Email.com ", "(11) 98765-4321", "São Paulo - SP"),
        ("JOÃO PEDRO alves", "joao.alves@email", "11 9 1234 5678", "Campinas - SP"),
        ("ana   clara", "ANA.CLARA@EMAIL.COM", "+55 (21) 99876-5432", "Rio de Janeiro - RJ"),
        ("carlos eduardo ", "carlos.edu@email.com", "31-98888-7777", "Belo Horizonte - MG"),
        ("maria   eduarda SANTOS", "maria.santos@email.com", "11987654321", "São Paulo - SP"),
        ("fernanda  LIMA", "fernanda.lima@email.com.br", "(41)99123-4567", "Curitiba - PR"),
        ("  rafael souza", "rafael souza@email.com", "21 99555 4444", "Niterói - RJ"),
    ]
    for i, linha in enumerate(brutos, start=5):
        for k, v in enumerate(linha, start=1):
            ws.cell(row=i, column=k, value=v)
        for col in "EFGHIJ":
            ws[f"{col}{i}"].fill = PatternFill("solid", fgColor=AZUL_CLARO)
            ws[f"{col}{i}"].border = borda
    ws["A14"] = "Gabarito (fórmulas em português, separador ponto e vírgula):"
    ws["A14"].font = Font(bold=True, color=AZUL_ESCURO)
    gabarito = [
        "E5  =PRI.MAIÚSCULA(ARRUMAR(A5))",
        "F5  =MINÚSCULA(ARRUMAR(B5))",
        "G5  =CONCAT(SEERRO(EXT.TEXTO(C5;SEQUÊNCIA(NÚM.CARACT(C5));1)*1;\"\"))   (Excel 365)",
        "H5  =ARRUMAR(TEXTOANTES(D5;\"-\"))      I5  =ARRUMAR(TEXTODEPOIS(D5;\"-\"))   (Excel 365)",
        "J5  =E(ÉNÚM(LOCALIZAR(\"@\";F5));ÉNÚM(LOCALIZAR(\".\";F5;LOCALIZAR(\"@\";F5)));NÃO(ÉNÚM(LOCALIZAR(\" \";F5))))",
        "Depois: copie, Colar Especial > Valores, e use Dados > Remover Duplicatas pelo e-mail limpo (linhas 5 e 9 são a mesma pessoa).",
    ]
    for j, txt in enumerate(gabarito, start=15):
        ws[f"A{j}"] = txt
        ws[f"A{j}"].font = Font(name="Consolas", size=10)
    larguras(ws, {"A": 30, "B": 30, "C": 22, "D": 24, "E": 24, "F": 28, "G": 22, "H": 18, "I": 8, "J": 22})


def aba_scorecard(wb: Workbook) -> None:
    ws = wb.create_sheet("Scorecard")
    titulo(ws, "Scorecard de entrevista", "Notas de 1 a 5 por competência. A nota final é uma média ponderada (SOMARPRODUTO).")
    ws["A4"] = "Competência"
    ws["B4"] = "Peso"
    cands = ["Candidato A", "Candidato B", "Candidato C"]
    for k, c in enumerate(cands):
        ws.cell(row=4, column=3 + k, value=c)
    comps = [("Conhecimento técnico", 0.30), ("Comunicação", 0.20), ("Resolução de problemas", 0.25),
             ("Fit com valores", 0.15), ("Experiência relevante", 0.10)]
    for i, (comp, peso) in enumerate(comps, start=5):
        ws[f"A{i}"] = comp
        ws[f"B{i}"] = peso
        ws[f"B{i}"].number_format = "0%"
        for k in range(len(cands)):
            ws.cell(row=i, column=3 + k, value=random.randint(2, 5))
    fim = 4 + len(comps)
    ws[f"A{fim + 1}"] = "Soma dos pesos (precisa dar 100%)"
    ws[f"B{fim + 1}"] = f"=SUM(B5:B{fim})"
    ws[f"B{fim + 1}"].number_format = "0%"
    ws[f"A{fim + 2}"] = "Nota final ponderada"
    ws[f"A{fim + 3}"] = "Recomendação"
    for k in range(len(cands)):
        col = get_column_letter(3 + k)
        ws[f"{col}{fim + 2}"] = f"=SUMPRODUCT($B$5:$B${fim},{col}5:{col}{fim})"
        ws[f"{col}{fim + 2}"].number_format = "0.00"
        ws[f"{col}{fim + 2}"].font = Font(bold=True, color=AZUL_ESCURO)
        ws[f"{col}{fim + 3}"] = (f'=IF({col}{fim + 2}>=4,"Avançar",IF({col}{fim + 2}>=3,'
                                 f'"Avaliar com gestor","Não avançar"))')
    for col in range(1, 3 + len(cands)):
        cel = ws.cell(row=4, column=col)
        cel.font = Font(bold=True, color=BRANCO)
        cel.fill = PatternFill("solid", fgColor=AZUL)
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5")
    dv.error = "Use notas inteiras de 1 a 5."
    dv.errorTitle = "Nota inválida"
    ws.add_data_validation(dv)
    dv.add(f"C5:{get_column_letter(2 + len(cands))}{fim}")
    ws.conditional_formatting.add(f"C5:{get_column_letter(2 + len(cands))}{fim}", CellIsRule(
        operator="lessThanOrEqual", formula=["2"], fill=PatternFill("solid", fgColor="FECACA")))
    larguras(ws, {"A": 36, "B": 10, "C": 15, "D": 15, "E": 15})


def aba_exercicios(wb: Workbook) -> None:
    ws = wb.create_sheet("Exercícios")
    titulo(ws, "Exercícios guiados", "Faça na ordem. As respostas estão no módulo Exercícios do tutorial.")
    tarefas = [
        "1. Na aba Candidatos, filtre só os candidatos 'Em andamento' do recrutador 'Ana Lima'.",
        "2. Classifique a tabela por 'Dias parado', do maior para o menor. Quem precisa de retorno hoje?",
        "3. Crie uma Segmentação de Dados (Tabela > Inserir Segmentação) para a coluna Fonte.",
        "4. Insira uma Tabela Dinâmica: linhas = Cargo, colunas = Etapa, valores = Contagem de ID.",
        "5. No Painel, crie um gráfico de Funil com o bloco 'Funil'.",
        "6. Em uma célula vazia, use PROCX para trazer o e-mail do candidato CAND-010.",
        "7. Use FILTRO para listar os nomes com Nota Entrevista >= 4,5 (Excel 365).",
        "8. Adicione um novo candidato na primeira linha vazia abaixo da tabela e veja as fórmulas e listas se estenderem.",
        "9. Na aba Dados_Brutos, preencha as colunas azuis com fórmulas de limpeza.",
        "10. Proteja a aba Painel (Revisão > Proteger Planilha) para ninguém apagar as fórmulas.",
    ]
    for i, t in enumerate(tarefas, start=4):
        ws[f"A{i}"] = t
        ws[f"B{i}"] = "☐"
        ws[f"B{i}"].alignment = Alignment(horizontal="center")
    dv = DataValidation(type="list", formula1='"☐,☑"')
    ws.add_data_validation(dv)
    dv.add(f"B4:B{3 + len(tarefas)}")
    larguras(ws, {"A": 110, "B": 8})


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)
    cands = gerar_candidatos()
    aba_listas(wb)
    aba_vagas(wb)
    ultima = aba_candidatos(wb, cands)
    aba_painel(wb, ultima)
    aba_brutos(wb)
    aba_scorecard(wb)
    aba_exercicios(wb)
    ordem = ["Painel", "Candidatos", "Vagas", "Scorecard", "Dados_Brutos", "Exercícios", "Listas"]
    wb._sheets = [wb[n] for n in ordem]  # noqa: SLF001
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = AZUL if ws.title in ("Painel", "Candidatos", "Vagas") else "93C5FD"
    wb.active = 0
    saida = Path(__file__).resolve().parent.parent / "arquivos" / "planilha-pratica-recrutamento.xlsx"
    saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(saida)
    print(f"ok: {saida}")


if __name__ == "__main__":
    main()
