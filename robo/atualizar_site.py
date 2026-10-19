"""Calcula o IPA a partir do relatório do GisWorks e cria a aba nova do site
com a análise completa (mesmas seções do doc do IPA).

Uso: python robo/atualizar_site.py relatorio.xlsx [abertas.xlsx]

Regras (as mesmas das análises manuais):
- base = OS ("Ordem de Serviço") fechadas no mês, exceto origem Obra / reclamante "Obra"
  (OS de vias fechadas por equipes de obra continuam na base);
- inspeções fora do IPA; OS não executada conta;
- no prazo = campo "Fechada no Prazo" = "Sim"; IPA com 2 casas;
- prazo de atendimento: 24h vias especiais, 48h demais (Anexo 2, item 4.1.3.8).

O robô mantém UMA aba automática por dia: a cada execução substitui a aba
automática do mesmo dia. Abas feitas à mão (sem "auto": true) nunca são tocadas.
"""
import calendar
import json
import math
import sys
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
HIST = RAIZ / "historico"
TZ = ZoneInfo("America/Manaus")

# faixas de nota do IPA (Anexo 3, Tabela 3-4). Ajuste aqui se a tabela tiver mais faixas.
FAIXAS_NOTA = [(95.0, "1,0"), (90.0, "0,8"), (0.0, "0,0")]
ORIGENS_RECLAMACAO = {"Sistema - Criada no Call Center", "App - Android", "App - iOS", "Web"}
METAS = {"atrasos_dia": 3, "fila_h": 8, "sem_exec": 12, "erros": 20}

GRUPOS_MOTIVO = [
    ("Escopo / triagem", ["transformar em obra", "encontrado normal", "não localizado", "nao localizado",
                          "outros", "areas privadas", "áreas privadas", "modernização", "terminais"]),
    ("Despacho / equipe", ["veiculo improprio", "veículo impróprio", "luminária correta", "luminaria correta"]),
    ("Acesso / terceiros", ["terceiro", "acesso restrito", "área de risco", "area de risco", "tensão", "tensao"]),
    ("Furto", ["furt", "vandalismo"]),
    ("Rede / concessionária AME", ["fornecimento de energia", "ame"]),
]


# ---------------------------------------------------------------- utilidades
def br(n, casas=0):
    s = f"{n:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pc(n, casas=2):
    return f"{br(n, casas)}%"


def pct(sim, total):
    return 100.0 * sim / total if total else 0.0


def nota(ipa):
    for limite, n in FAIXAS_NOTA:
        if round(ipa, 2) >= limite:
            return n
    return FAIXAS_NOTA[-1][1]


def horas(td):
    return td.dt.total_seconds() / 3600


def hh(v):
    return "—" if v is None or pd.isna(v) else f"{br(v, 1)} h"


def data_br(s):
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def eh_especial(p):
    p = str(p).lower()
    return "especia" in p and "não" not in p and "nao" not in p


def turno_de(eq):
    e = str(eq).strip().upper()
    if e.startswith("TURNO"):
        return "Turno " + e.split()[1].split("-")[0] if len(e.split()) > 1 else "Turno"
    if "OBRA" in e:
        return "Equipes de obra (OS de vias)"
    return "Equipes de apoio e outras"


def grupo_motivo(m):
    t = str(m).strip().lower()
    for g, chaves in GRUPOS_MOTIVO:
        if any(c in t for c in chaves):
            return g
    return "Outros"


def tabela(cab, linhas):
    out = ["| " + " | ".join(cab) + " |", "| " + " | ".join("---" for _ in cab) + " |"]
    out += ["| " + " | ".join(str(c) for c in l) + " |" for l in linhas]
    return "\n".join(out)


def agrega(d, col):
    g = d.groupby(col).agg(os=("prazo", "size"), sim=("prazo", "sum"))
    g["atrasos"] = g["os"] - g["sim"]
    g["ipa"] = 100 * g["sim"] / g["os"]
    return g


def lista_e(itens):
    itens = list(itens)
    return itens[0] if len(itens) == 1 else ", ".join(itens[:-1]) + " e " + itens[-1]


def parse_atualizado(s):
    try:
        return datetime.strptime(s, "%d/%m/%Y %Hh%M")
    except Exception:
        return None


# ---------------------------------------------------------------- leitura
def ler_relatorio(arquivo):
    df = pd.read_excel(arquivo, sheet_name="Relatório", header=1)
    df.columns = [str(c).strip() for c in df.columns]
    for c, n in [("Data", "ab"), ("Data Fechamento", "fech"), ("Início da Execução", "iexec"), ("Data Lote", "lote")]:
        df[n] = data_br(df[c]) if c in df else pd.NaT
    try:
        mat = pd.read_excel(arquivo, sheet_name="Material", header=1)
        mat.columns = [str(c).strip() for c in mat.columns]
    except Exception:
        mat = pd.DataFrame()
    return df, mat


def ler_abertas(arquivo):
    if not arquivo or not Path(arquivo).exists():
        return None
    try:
        brut = pd.read_excel(arquivo, header=None)
        # cabeçalho = primeira linha que tenha "ID" e "Data"
        def cab(i):
            c = {str(x).strip() for x in brut.iloc[i].tolist()}
            return "Data" in c and ("ID" in c or "Número da OS" in c)
        lin = next(i for i in range(min(10, len(brut))) if cab(i))
        ab = pd.read_excel(arquivo, header=lin)
        ab.columns = [str(c).strip() for c in ab.columns]
        ab["ab"] = data_br(ab["Data"])
        return ab
    except Exception as e:
        msg = f"Aviso: não consegui ler as OS em aberto ({e})"
        try:
            brut = pd.read_excel(arquivo, header=None, nrows=4)
            msg += "\nPrimeiras linhas: " + " || ".join(" | ".join(map(str, r)) for r in brut.values.tolist())
        except Exception:
            pass
        print(msg)
        with open(Path(__file__).parent / "diagnostico.txt", "a", encoding="utf-8") as f:
            f.write(msg + "\n")
        return None


def sem_obra_mask(d):
    o = d["Origem"].astype(str).str.strip().eq("Obra") if "Origem" in d else False
    r = d["Reclamante"].astype(str).str.strip().str.upper().eq("OBRA") if "Reclamante" in d else False
    return ~(o | r)


# ---------------------------------------------------------------- análise
def analisar(df, mat, abertas, agora, ant):
    ini_mes = datetime(agora.year, agora.month, 1)
    dfm = df[(df["fech"] >= ini_mes) & (df["fech"] <= agora)].copy()
    os_ = dfm[dfm["Tipo"] == "Ordem de Serviço"]
    base = os_[sem_obra_mask(os_)].copy()
    if base.empty:
        raise SystemExit("Relatório sem OS no período — nada a publicar.")
    base["prazo"] = base["Fechada no Prazo"] == "Sim"
    base["esp"] = base["Prioridade"].map(eh_especial)
    base["venc"] = base["ab"] + pd.to_timedelta(base["esp"].map({True: 24, False: 48}), unit="h")
    base["exec"] = base["Ocorrência"].astype(str).str.upper().eq("EXECUTADO")
    base["turno"] = base["Equipe"].map(turno_de)
    base["equipe"] = base["Equipe"].astype(str).str.strip()

    R = {}  # resultados numéricos (alguns vão para o json da aba)
    total, no_prazo = len(base), int(base["prazo"].sum())
    fora = total - no_prazo
    ipa = round(pct(no_prazo, total), 2)
    ultimo = base["fech"].max()
    atr = base[~base["prazo"]]
    dias_dec = max((agora - ini_mes).total_seconds() / 86400, 0.5)
    dias_mes = calendar.monthrange(agora.year, agora.month)[1]

    # gráfico
    dias = []
    pdia = base.groupby(base["fech"].dt.day)
    for dia in range(1, ultimo.day + 1):
        g = pdia.get_group(dia) if dia in pdia.groups else base.iloc[0:0]
        dias.append({"dia": f"{dia:02d}", "os": len(g), "atrasos": int((~g["prazo"]).sum())})

    mesmo_mes = ant.get("mes") == f"{agora:%m/%Y}"
    ant_ex = ant.get("extra", {}) if mesmo_mes else {}
    t_ant = parse_atualizado(ant.get("atualizado_em", "")) if mesmo_mes else None

    S = []  # seções em markdown
    n_obra = int((~sem_obra_mask(os_)).sum())
    obra_eq = int(base["equipe"].str.upper().str.contains("OBRA").sum())
    n_insp = int((dfm["Tipo"] == "Inspeção").sum())
    n_exec = int(base["exec"].sum())
    S.append(
        f"**Período considerado:** fechamentos de 01/{agora:%m/%Y} 00h00 até {ultimo:%d/%m/%Y} às {ultimo:%Hh%M} "
        f"(último fechamento no relatório, gerado automaticamente às {agora:%Hh%M}). Base: todas as OS fechadas no "
        f"período, exceto as de origem Obra ({br(n_obra)} OS, reclamante \"Obra\")"
        + (f"; as {br(obra_eq)} OS de vias fechadas por equipes de obra estão incluídas" if obra_eq else "")
        + f". Das {br(total)}, {br(n_exec)} foram executadas e {br(total - n_exec)} fechadas sem execução. "
          f"As inspeções ({br(n_insp)}) não entram no IPA. O total confere com o sistema "
          f"({br(no_prazo)} + {br(fora)} = {br(total)}).")
    S.append("<!-- GRAFICO -->")

    # ---------- blocos de cálculo ----------
    # novos desde a aba anterior
    novos = base[base["fech"] > t_ant] if t_ant else None
    # ritmo dos últimos 10 dias
    ini10 = max(ini_mes, datetime(agora.year, agora.month, agora.day) - timedelta(days=9))
    u10 = base[base["fech"] >= ini10]
    dias10 = max((agora - ini10).total_seconds() / 86400, 0.5)
    ritmo = (len(u10) - int(u10["prazo"].sum())) / dias10
    R["ritmo"] = round(ritmo, 1)

    # causa raiz
    ini_m = pd.Timestamp(ini_mes)
    passivo_v = atr["venc"] < ini_m
    fila_v = ~passivo_v & (atr["iexec"].isna() | (atr["iexec"] > atr["venc"]))
    campo_v = ~passivo_v & ~fila_v
    ab_antes = int((atr["ab"] < ini_m).sum())
    mais7 = int(((atr["fech"] - atr["ab"]) > pd.Timedelta(days=7)).sum())
    noite = int((atr["ab"].dt.hour >= 18).sum())
    com_lote = base[base["lote"].notna()].copy()
    com_lote["h_fila"] = horas(com_lote["lote"] - com_lote["ab"]).clip(lower=0)
    com_lote["h_campo"] = horas(com_lote["fech"] - com_lote["lote"]).clip(lower=0)
    com_lote["h_total"] = horas(com_lote["fech"] - com_lote["ab"]).clip(lower=0)
    med = {}
    for esp in (True, False):
        for pz in (True, False):
            g = com_lote[(com_lote["esp"] == esp) & (com_lote["prazo"] == pz)]
            med[(esp, pz)] = (len(g), g["h_fila"].median(), g["h_campo"].median(), g["h_total"].median())
    R["fila_demais"] = None if pd.isna(med[(False, True)][1]) else round(float(med[(False, True)][1]), 1)

    # prioridade
    pri = agrega(base, "esp")
    ipa_esp = float(pri.loc[True, "ipa"]) if True in pri.index else None

    # logradouro concentrado
    conc = None
    if "Logradouro" in base and len(atr):
        lg = atr.groupby(atr["Logradouro"].astype(str).str.strip()).size().sort_values(ascending=False)
        if len(lg) and lg.iloc[0] >= max(15, 0.08 * fora):
            nome_lg = lg.index[0]
            n_lg = base[base["Logradouro"].astype(str).str.strip() == nome_lg]
            sem = base.drop(n_lg.index)
            conc = (nome_lg, int(lg.iloc[0]), len(n_lg), pct(sem["prazo"].sum(), len(sem)))

    # não executadas e motivos
    nexe = base[~base["exec"]].copy()
    nexe["motivo"] = nexe["Motivos não realização"].fillna("").astype(str).str.strip().replace({"": "Sem motivo", "nan": "Sem motivo"})
    nexe["grupo"] = nexe["motivo"].map(grupo_motivo)
    mot = nexe.groupby("motivo").agg(os=("prazo", "size"), atr=("prazo", lambda s: int((~s).sum())))

    def mot_n(chave):
        m = mot[mot.index.str.lower().str.contains(chave)]
        return int(m["os"].sum()), int(m["atr"].sum())

    vi_os, vi_atr = mot_n("veiculo improprio|veículo impróprio")
    lc_os, lc_atr = mot_n("luminária correta|luminaria correta")
    ame_os, _ = mot_n("fornecimento de energia")
    furto_os = int(nexe[nexe["grupo"] == "Furto"].shape[0])
    acesso_os, _ = mot_n("acesso restrito")
    R["sem_exec"] = round(pct(len(nexe), total), 1)

    # fechadas mais de uma vez
    multi = base.sort_values("fech").groupby("ID")
    tam = multi.size()
    ids_multi = tam[tam > 1].index
    bm = base[base["ID"].isin(ids_multi)].sort_values("fech")
    primeiro = bm.groupby("ID").first()
    redesp_vi = int(primeiro["Motivos não realização"].astype(str).str.lower()
                    .str.contains("veiculo improprio|veículo impróprio").sum()) if len(primeiro) else 0
    muda = int((bm.groupby("ID")["prazo"].nunique() > 1).sum()) if len(bm) else 0
    unica = base.sort_values("fech").drop_duplicates("ID", keep="last")

    # equipes
    eq = agrega(base, "equipe")
    eq["dias"] = base.groupby("equipe")["fech"].apply(lambda s: s.dt.date.nunique())
    eq["nexe"] = 100 * base.groupby("equipe")["exec"].apply(lambda s: (~s).sum()) / eq["os"]
    piores = eq[(eq["os"] >= 20) & (eq["atrasos"] > 0)].sort_values("ipa")
    R["equipes"] = {k: [int(r.os), int(r.atrasos)] for k, r in eq.iterrows()}

    # materiais e erros de lançamento
    M = materiais(dfm, mat)

    # backlog
    B = backlog(abertas, agora) if abertas is not None else None
    if B:
        R.update({k: B[k] for k in ("abertas", "vencidas", "venc_esp", "sem_equipe", "venc_sem_equipe")})

    # projeção / orçamento
    os_dia = total / dias_dec
    mes_total = os_dia * dias_mes
    vencidas = B["vencidas"] if B else 0
    rest = dias_mes - agora.day + 1
    orc = []
    for meta, frac, nt in ((90, 0.10, "0,8"), (95, 0.05, "1,0")):
        perm = math.floor(mes_total * frac)
        saldo = perm - fora - vencidas
        orc.append((meta, nt, perm, saldo, saldo / rest))
    p95 = max(0, math.ceil(0.95 * total - no_prazo - 1e-9))

    # ---------- RESUMO EXECUTIVO ----------
    S.append("## Resumo executivo")
    var = ""
    if mesmo_mes and ant.get("ipa") is not None:
        dv = ipa - ant["ipa"]
        nv = ""
        if novos is not None and len(novos):
            na = int((~novos["prazo"]).sum())
            nv = (f"entraram {br(len(novos))} fechamentos com {br(na)} atrasos "
                  f"({pc(pct(len(novos) - na, len(novos)))} no prazo), e ")
        var = (f" Desde a aba anterior ({ant['atualizado_em']}) {nv}o índice "
               + ("ficou estável." if abs(dv) < 0.005 else
                  f"{'subiu' if dv > 0 else 'caiu'} {'+' if dv > 0 else '−'}{br(abs(dv), 2)} p.p. "
                  f"(de {pc(ant['ipa'])})."))
    S.append(f"O IPA de {agora:%m/%Y} está em **{pc(ipa)} — nota {nota(ipa)}** "
             f"({br(no_prazo)} de {br(total)} OS no prazo).{var}"
             + (f" Para estar em 95,00% hoje seriam necessários {br(p95)} atrasos a menos." if ipa < 95 else ""))

    af = med[(True, False)], med[(False, False)]
    filas = [x[1] for x in af if x[0] and not pd.isna(x[1])]
    campos = [x[2] for x in af if x[0] and not pd.isna(x[2])]
    if filas and int(fila_v.sum()) >= int(campo_v.sum()):
        S.append(f"O atraso é de fila: {br(int(fila_v.sum()))} dos {br(fora)} atrasos venceram antes de a equipe "
                 f"iniciar a execução. As OS atrasadas esperaram de {br(min(filas), 0)} a {br(max(filas), 0)} horas "
                 f"(mediana) até entrar num lote e levaram cerca de {br(max(campos), 0)} horas do lote ao fechamento.")
    bl = []
    if conc:
        bl.append(f"- **Concentração em {conc[0].title()}:** {br(conc[1])} atrasos ({pc(pct(conc[1], fora), 0)} do total). "
                  f"Sem esse logradouro, o IPA estaria em {pc(conc[3])}.")
    bl.append(f"- **Não executadas:** {br(len(nexe))} OS ({pc(pct(len(nexe), total), 0)}) fechadas sem execução"
              + (f"; veículo impróprio ({br(vi_os)}) e voltar com luminária correta ({br(lc_os)}) geram "
                 f"{br(vi_atr + lc_atr)} atrasos." if vi_os + lc_os else "."))
    bl.append(f"- **Passivo e noite:** {br(ab_antes)} atrasos são OS abertas antes de 01/{agora:%m} e "
              f"{br(noite)} foram abertas entre 18h e 24h.")
    if len(piores):
        p3 = piores.head(3)
        bl.append(f"- **Equipes:** {lista_e(f'{k} ({pc(r.ipa)})' for k, r in p3.iterrows())} somam "
                  f"{br(int(p3['atrasos'].sum()))} dos {br(fora)} atrasos.")
    S.append("\n".join(bl))
    if B and B["vencidas"]:
        cmp = ""
        if "vencidas" in ant_ex:
            a0 = ant_ex["vencidas"]
            cmp = (f" {'subiram' if B['vencidas'] > a0 else 'caíram' if B['vencidas'] < a0 else 'seguem'} "
                   f"{'de ' + br(a0) + ' para ' if B['vencidas'] != a0 else 'em '}{br(B['vencidas'])}")
        else:
            cmp = f" são {br(B['vencidas'])}"
        S.append(f"**Alerta do dia — vencidas em aberto:** as OS de vias vencidas em aberto{cmp}"
                 f" ({br(B['venc_esp'])} de via especial). {br(B['venc_sem_equipe'])} delas estão sem equipe e "
                 f"{br(B['venc_7d'])} venceram há mais de 7 dias"
                 + (f"; o {B['pior_eq'][0]} está com {br(B['pior_eq'][1])} vencidas" if B["pior_eq"] else "")
                 + ". Cada uma vira atraso quando for fechada.")
    o90, o95 = orc
    if ipa >= 95 and o95[3] >= 0:
        meta_txt = (f"≥95,00% (nota 1,0) está sendo cumprido; o mês comporta mais {br(o95[3])} atrasos "
                    f"(cerca de {br(o95[4], 1)} por dia) contra um ritmo de {br(ritmo, 1)} por dia nos últimos dias.")
    else:
        meta_txt = (f"≥90,00% {'está garantida' if o90[3] >= 0 and o90[4] >= ritmo else 'está em risco'} no ritmo atual "
                    f"(saldo de {br(o90[3])} atrasos até {dias_mes}/{agora:%m}). ")
        if o95[3] > 0:
            meta_txt += (f"Para 95,00% o saldo é de {br(o95[3])} atrasos novos (cerca de {br(o95[4], 1)} por dia), "
                         f"contra um ritmo de cerca de {br(ritmo, 1)} por dia nos últimos 10 dias"
                         + (", considerando as vencidas em aberto como atrasos futuros." if B else "."))
        else:
            meta_txt += (f"95,00% ficou fora de alcance este mês: "
                         + (f"com as {br(vencidas)} vencidas, " if vencidas else "")
                         + f"já são {br(fora + vencidas)} atrasos para um limite de cerca de {br(o95[2])}. "
                         f"O foco passa a ser zerar as vencidas e chegar ao próximo mês sem passivo.")
    S.append(f"**Meta:** {meta_txt}")

    # ---------- PLANO DE AÇÃO ----------
    S.append("## Plano de ação")
    hoje_s, d1, d2, d5 = (f"{(agora + timedelta(days=k)):%d/%m}" for k in (0, 1, 2, 5))
    acoes = []
    if B and B["sem_equipe"]:
        acoes.append((f"Despachar na abertura do próximo turno as {br(B['sem_equipe'])} OS de vias sem equipe "
                      f"({br(B['venc_sem_equipe'])} vencidas, {br(B['esp_sem_equipe'])} de via especial)",
                      "CCO / Despacho", hoje_s, "Nenhuma OS sem equipe com mais de 8h"))
    if B and B["pior_eq"] and B["pior_eq"][1] >= 3:
        acoes.append((f"Fechar ou redistribuir as {br(B['pior_eq'][1])} OS vencidas que estão com o {B['pior_eq'][0]}",
                      "CCO / Supervisores", hoje_s, f"Zero vencida com o {B['pior_eq'][0]}"))
    if ipa_esp is not None and ipa_esp < 95:
        acoes.append(("Via especial (24h) não espera lote: despacho imediato à equipe mais próxima",
                      "CCO / Despacho", d1, f"Vias especiais de {pc(ipa_esp)} para mais de 95,00%"))
    if vi_os >= 10:
        acoes.append((f"Conferir o veículo antes do despacho: veículo impróprio gerou {br(vi_os)} não execuções, "
                      f"{br(vi_atr)} atrasos e {br(redesp_vi)} redespachos",
                      "CCO / Despacho", d1, "Redespachos por veículo impróprio pela metade"))
    pouco = eq[eq.index.str.upper().str.startswith("TURNO") & (eq["dias"] < 0.5 * agora.day)].sort_values("dias")
    if len(pouco):
        pz = pouco.head(3)
        acoes.append((f"Escalar {lista_e(pz.index)} nos dias úteis "
                      f"({lista_e(str(int(d)) for d in pz['dias'])} dias ativos no mês)",
                      "Coordenação", d1, f"+{len(pz)} equipes por dia"))
    if M and len(M["equipes"]):
        top = M["equipes"].head(2)
        partes = []
        for k, r in top.iterrows():
            col = max(M["regras"], key=lambda c: r[c])
            partes.append(f"do {k} ({br(int(r[col]))} {M['curto'][col]})")
        acoes.append((f"Corrigir o lançamento {lista_e(partes)}", "Supervisores", d2, "Serviço e material coerentes"))
    if ame_os + furto_os + acesso_os:
        acoes.append((f"Protocolar justificativa dos casos de falta de energia da AME ({br(ame_os)}), furto "
                      f"({br(furto_os)}) e acesso restrito ({br(acesso_os)}), com foto (item 4.1.3.8)",
                      "Gestão contratual", d2, f"Até {br(ame_os + furto_os + acesso_os)} OS fora do SMD"))
    if len(piores) and piores.iloc[0]["ipa"] < 85:
        k, r = piores.index[0], piores.iloc[0]
        acoes.append((f"Acompanhar o {k} (IPA {pc(r.ipa)}, {br(int(r.atrasos))} atrasos): revisar rota e lotes com o supervisor",
                      "Supervisores", d5, f"{k} acima de 90,00%"))
    S.append(tabela(["#", "Ação", "Responsável", "Prazo", "Resultado esperado"],
                    [(i + 1, *a) for i, a in enumerate(acoes)]))

    # ---------- IPA COM RECORTES ----------
    S.append("## IPA com recortes")
    sem_c = None
    if conc:
        sem_c = base[base["Logradouro"].astype(str).str.strip() != conc[0]]
    rec = base[base["Origem"].isin(ORIGENS_RECLAMACAO)]

    def lr(nome, d, p=None):
        s = int(d["prazo"].sum()) if p is None else p
        v = pct(s, len(d))
        return (nome, br(len(d)), pc(v), nota(v))

    linhas = [lr("Todas as OS, exceto origem Obra (base adotada)", base),
              lr("Cada OS uma vez, pelo último fechamento", unica)]
    if sem_c is not None:
        linhas.append(lr(f"Sem {conc[0].title()} ({br(conc[2])} OS, {br(conc[1])} atrasos)", sem_c))
    linhas.append(lr("Só reclamações de usuários (Call Center, App, Web)", rec))
    linhas.append(lr("Leitura estrita: só OS executadas e no prazo", base, int((base["prazo"] & base["exec"]).sum())))
    S.append(tabela([f"Recorte (01/{agora:%m}–{agora:%d/%m})", "OS", "No prazo", "Nota IPA"], linhas))

    def it(nome, r):
        return f"{nome} {pc(r.ipa)} ({br(int(r.atrasos))} atrasos em {br(int(r.os))})"

    bl = []
    pr = [it("vias especiais" if k else "demais vias", r) for k, r in pri.sort_index(ascending=False).iterrows()]
    if "esp" in ant_ex:
        pr_ant = ant_ex["esp"]
        bl.append(f"- **Prioridade:** {'; '.join(pr)}. Na aba anterior: {pc(pr_ant[0])} e {pc(pr_ant[1])}.")
    else:
        bl.append(f"- **Prioridade:** {'; '.join(pr)}.")
    R["esp"] = [round(float(pri.loc[True, "ipa"]), 2) if True in pri.index else None,
                round(float(pri.loc[False, "ipa"]), 2) if False in pri.index else None]
    ori = agrega(base, "Origem").sort_values("ipa", ascending=False)
    bl.append("- **Origem:** " + "; ".join(
        f"{k} {pc(r.ipa)}" + (f" ({br(int(r.atrasos))} atrasos em {br(int(r.os))})" if r.ipa < 95 else f" ({br(int(r.os))})")
        for k, r in ori.iterrows()) + ".")
    exe = agrega(base, "exec")
    bl.append("- **Executada x não executada:** " + "; ".join(
        it("executadas" if k else "não executadas", r) for k, r in exe.sort_index(ascending=False).iterrows()) + ".")
    zon = agrega(base[base["Área"].astype(str).str.upper().str.startswith("ZONA")], "Área").sort_values("atrasos", ascending=False)
    bl.append("- **Zonas:** " + "; ".join(it(str(k).title(), r) for k, r in zon.iterrows()) + ".")
    S.append("\n".join(bl))

    # ---------- CAUSA RAIZ ----------
    S.append("## Causa raiz dos atrasos: fila x campo")
    S.append(f"Dos {br(fora)} atrasos, {br(int(fila_v.sum()))} venceram na fila antes de a equipe iniciar a execução, "
             f"{br(int(passivo_v.sum()))} já estavam vencidos antes de 01/{agora:%m} (passivo) e "
             f"{br(int(campo_v.sum()))} venceram em campo. Mediana de horas por etapa "
             f"({br(len(com_lote))} OS com lote; ronda, telegestão e parte das demais fecham sem lote):")
    nomes = {(True, True): "Via especial – no prazo", (True, False): "Via especial – atrasada",
             (False, True): "Demais vias – no prazo", (False, False): "Demais vias – atrasada"}
    S.append(tabela([f"Situação (01/{agora:%m}–{agora:%d/%m})", "OS com lote", "Abertura → lote (fila)",
                     "Lote → fechamento (campo)", "Total"],
                    [(nomes[k], br(v[0]), hh(v[1]), hh(v[2]), hh(v[3])) for k, v in med.items()]))
    txt = f"O passivo pesa: {br(ab_antes)} atrasos são OS abertas antes de 01/{agora:%m} e {br(mais7)} levaram mais de 7 dias para fechar."
    if "Logradouro" in base and len(atr):
        lg = atr.groupby(atr["Logradouro"].astype(str).str.strip().str.title()).size().sort_values(ascending=False)
        if conc:
            lg = lg[lg.index != conc[0].title()]
        top = lg.head(4)
        if len(top):
            txt += (f" {'Fora ' + conc[0].title() + ' (' + br(conc[1]) + '), os' if conc else 'Os'} logradouros com mais atrasos são "
                    + lista_e(f"{k} ({br(int(v))})" for k, v in top.items()) + ".")
    if novos is not None and len(novos) and (~novos["prazo"]).sum():
        nn = novos[~novos["prazo"]]
        fn = int((nn["iexec"].isna() | (nn["iexec"] > nn["venc"])).sum())
        txt += (f" Dos {br(len(nn))} atrasos novos desde a aba anterior, {br(fn)} vieram da fila "
                f"({lista_e(f'{v} pelo {k}' for k, v in nn['equipe'].value_counts().head(3).items())}).")
    S.append(txt)

    # ---------- TURNOS E EQUIPES ----------
    S.append("## Turnos e equipes")
    tu = agrega(base, "turno")
    ordem = sorted([t for t in tu.index if t.startswith("Turno")]) + \
            [t for t in ("Equipes de apoio e outras", "Equipes de obra (OS de vias)") if t in tu.index]
    tu = tu.reindex(ordem)
    tt = tu[tu.index.str.startswith("Turno")]
    if len(tt):
        S.append(f"O {tt['ipa'].idxmax()} é o mais pontual ({pc(tt['ipa'].max())}); o {tt['atrasos'].idxmax()} tem o maior "
                 f"número de atrasos ({br(int(tt['atrasos'].max()))}). O IPA é atribuído à equipe que fechou a OS.")
    S.append(tabela(["Turno", "OS", "IPA", "Atrasos"],
                    [(k, br(int(r.os)), pc(r.ipa), br(int(r.atrasos))) for k, r in tu.iterrows()]))
    lista = eq[(eq["os"] >= 10) & (eq["atrasos"] > 0)].sort_values("ipa")
    resto = eq.drop(lista.index)
    linhas = [(k, int(r.dias), br(int(r.os)), pc(r.ipa), br(int(r.atrasos)), pc(r.nexe, 0)) for k, r in lista.iterrows()]
    if len(resto):
        linhas.append(("Demais equipes (sem atraso ou menos de 10 OS)", "—", br(int(resto["os"].sum())),
                       pc(pct(resto["sim"].sum(), resto["os"].sum())), br(int(resto["atrasos"].sum())), "—"))
    S.append(tabela(["Equipe (piores primeiro)", "Dias ativos", "OS", "IPA", "Atrasos", "Não executadas"], linhas))
    p6 = piores.head(6)
    txt = (f"As {len(p6)} piores com volume ({lista_e(p6.index)}) somam {br(int(p6['atrasos'].sum()))} dos "
           f"{br(fora)} atrasos.") if len(p6) else ""
    if ant_ex.get("equipes"):
        mud = []
        for k, r in p6.iterrows():
            if k in ant_ex["equipes"]:
                o0, a0 = ant_ex["equipes"][k]
                i0 = pct(o0 - a0, o0)
                if abs(r.ipa - i0) >= 0.01:
                    mud.append(f"{k} {'subiu' if r.ipa > i0 else 'caiu'} de {pc(i0)} para {pc(r.ipa)}")
        if mud:
            txt += " Em relação à aba anterior: " + "; ".join(mud) + "."
    S.append(txt)

    S.append("### Motivos de não execução")
    gm = nexe.groupby("grupo")
    lin = []
    for g, d in sorted(gm, key=lambda x: -len(x[1])):
        vc = d["motivo"].str.replace(r"\s+", " ", regex=True).value_counts()
        lin.append((g, ", ".join(f"{m[:1].upper() + m[1:].lower()} ({n})" for m, n in vc.items()),
                    br(len(d)), br(int((~d["prazo"]).sum()))))
    lin.append(("**Total**", "", f"**{br(len(nexe))}**", f"**{br(int((~nexe['prazo']).sum()))}**"))
    S.append(tabela(["Grupo", "Motivos", "OS", "Fora do prazo"], lin))
    S.append(f"Veículo impróprio sozinho soma {br(vi_atr)} atrasos. As OS não executadas contam no indicador "
             f"porque houve deslocamento de equipe para verificar a viabilidade.")
    if len(ids_multi):
        S.append("### OS fechadas mais de uma vez")
        mot1 = primeiro["Motivos não realização"].dropna().astype(str).str.strip().value_counts().head(2)
        S.append(f"{br(len(ids_multi))} OS aparecem em mais de um fechamento no período ({br(len(bm))} linhas, "
                 f"{br(len(bm) - len(ids_multi))} fechamentos a mais)."
                 + (f" Os primeiros fechamentos mais comuns foram {lista_e(f'{m.lower()} ({n})' for m, n in mot1.items())}, "
                    f"seguidos de redespacho." if len(mot1) else "")
                 + f" Em {br(muda)} delas o prazo muda entre um fechamento e outro. Contando cada OS uma vez, "
                   f"pelo último fechamento, o IPA seria {pc(pct(unica['prazo'].sum(), len(unica)))}.")

    # ---------- MATERIAIS ----------
    if M:
        S.append("## Materiais aplicados e erros de lançamento")
        S.append(f"Nas OS e inspeções da base (sem origem Obra) foram {br(M['n_incl'])} lançamentos de inclusão e "
                 f"{br(M['n_ret'])} de retirada.")
        S.append(tabela(["Material", "Incluído", "Retirado", "Registros com o item"], M["grupos"]))
        if len(M["equipes"]):
            S.append("Registros com erro, pela regra de conferência (troca sem nenhum material; relé no serviço sem relé "
                     "lançado; relé incluído sem retirada; relé lançado sem serviço de relé; OS não executada com "
                     "material incluído), equipes com 80 registros ou mais:")
            S.append(tabela(["Equipe", "Registros", "Com erro"] + [M["nome"][c] for c in M["regras"]],
                            [(k, br(int(r.reg)), pc(r.perc, 1), *[br(int(r[c])) for c in M["regras"]])
                             for k, r in M["equipes"].iterrows()]))
        t = M["tot"]
        ant_err = f" (na aba anterior: {pc(ant_ex['erros'], 1)})" if "erros" in ant_ex else ""
        S.append(f"No geral, {br(t['erro'])} de {br(t['reg'])} registros ({pc(t['perc'], 1)}) têm alguma inconsistência: "
                 f"{br(t['troca'])} trocas sem material, {br(t['rele_srv'])} relés no serviço sem relé lançado, "
                 f"{br(t['rele_incl'])} relés incluídos sem retirada e {br(t['rele_sem_srv'])} relés lançados sem "
                 f"serviço de relé{ant_err}.")
        R["erros"] = round(t["perc"], 1)
        S.append("Outras conferências:\n\n" + "\n".join(f"- {x}" for x in M["outras"]))

    # ---------- BACKLOG ----------
    S.append(f"## Backlog atual e orçamento de atrasos (posição {agora:%d/%m}, {agora:%Hh%M})")
    if B:
        cmp = ""
        if "abertas" in ant_ex:
            cmp = (f" Na aba anterior ({ant['atualizado_em']}) eram {br(ant_ex['abertas'])} em aberto e "
                   f"{br(ant_ex['vencidas'])} vencidas.")
        S.append(f"Das {br(B['abertas'])} OS de vias em aberto ({br(B['esp'])} de via especial), {br(B['vencidas'])} já "
                 f"venceram — {br(B['venc_esp'])} de via especial, {br(B['venc_7d'])} vencidas há mais de 7 dias — e "
                 f"{br(B['vence24'])} vencem nas próximas 24h.{cmp}")
        S.append(tabela(["Com quem está", "Em aberto", "Já vencidas", "Vencem em 24h"], B["tabela"]))
        if B["idade"]:
            S.append(f"Das {br(B['venc_sem_equipe'])} vencidas sem equipe: {B['idade']}.")
    else:
        S.append("*O relatório de OS em aberto não foi baixado nesta execução; o orçamento abaixo não considera as vencidas em aberto.*")
    S.append(f"**Orçamento de atrasos para {agora:%m/%Y}.** No ritmo atual (cerca de {br(os_dia, 0)} OS fechadas por dia) "
             f"o mês fecha com cerca de {br(mes_total, 0)} OS. Já temos {br(fora)} atrasos"
             + (f" e as {br(vencidas)} vencidas virarão mais {br(vencidas)} — {br(fora + vencidas)} no total." if vencidas else "."))
    S.append(tabela([f"Meta de {agora:%m/%Y}", "Atrasos permitidos no mês", f"Saldo para {agora:%d}–{dias_mes}/{agora:%m}",
                     "Atrasos por dia permitidos"],
                    [(f"≥{m},00% (nota {nt})", br(perm), br(sal), br(max(pd_, 0), 1)) for m, nt, perm, sal, pd_ in orc]))
    S.append(f"Nos últimos 10 dias a operação fez {br(len(u10) - int(u10['prazo'].sum()))} atrasos em {br(len(u10))} OS "
             f"({pc(pct(len(u10) - u10['prazo'].sum(), len(u10)))}, cerca de {br(ritmo, 1)} por dia).")
    if B and B["obras"]:
        S.append(f"Obras e cadastro: {br(B['obras'])} OS abertas ({br(B['obras90'])} com mais de 90 dias, "
                 f"{br(B['obras180'])} com mais de 180). Não entram no IPA.")

    # ---------- METAS ----------
    S.append("## Metas e premissas")
    def ant_v(chave, fmt):
        return fmt(ant_ex[chave]) if chave in ant_ex and ant_ex[chave] is not None else "—"
    ipa_ant = f"{pc(ant['ipa'])} ({br(ant['no_prazo'])} de {br(ant['os_fechadas'])})" if mesmo_mes and ant.get("ipa") is not None else "—"
    linhas = [("IPA acumulado do mês", f"{pc(ipa)} ({br(no_prazo)} de {br(total)})", ipa_ant, "≥ 90,00% (buscar 95,00%)"),
              ("Atrasos novos por dia (últimos 10 dias)", f"cerca de {br(ritmo, 1)}", ant_v("ritmo", lambda v: f"cerca de {br(v, 1)}"),
               f"≤ {METAS['atrasos_dia']}")]
    if B:
        linhas.append(("OS de vias vencidas em aberto", f"{br(B['vencidas'])} ({br(B['venc_esp'])} de via especial)",
                       ant_v("vencidas", br), "0"))
        linhas.append(("OS de vias sem equipe", f"{br(B['sem_equipe'])} ({br(B['venc_sem_equipe'])} vencidas)",
                       ant_v("sem_equipe", br), "0 com mais de 8h"))
    linhas.append(("Fila mediana (abertura → lote) das demais vias", hh(R["fila_demais"]), ant_v("fila_demais", hh),
                   f"< {METAS['fila_h']} h"))
    linhas.append(("OS fechadas sem execução", pc(R["sem_exec"], 0), ant_v("sem_exec", lambda v: pc(v, 0)), f"< {METAS['sem_exec']}%"))
    if M:
        linhas.append(("Registros com erro de lançamento", pc(R["erros"], 1), ant_v("erros", lambda v: pc(v, 1)),
                       f"< {METAS['erros']}%"))
    S.append(tabela(["Indicador", f"Hoje (01/{agora:%m}–{agora:%d/%m})", "Aba anterior", "Meta"], linhas))
    S.append("**Premissas**")
    S.append("\n".join([
        f"- Base: Relatório de Ordens Executadas e Inspeções/Materiais (fechamentos de 01/{agora:%m} a {agora:%d/%m/%Y}, "
        f"gerado às {agora:%Hh%M})" + (" e Relatório de Todas as Ordens em Aberto ou em Andamento" if B else "")
        + ", Manaus, GisWorks.",
        "- IPA com todas as OS fechadas no período, exceto origem Obra, pelo campo \"Fechada no Prazo\" e pela data de "
        "fechamento. Inspeções ficam fora. OS não executada conta no indicador.",
        "- Prazos e notas: Anexo 2 – Caderno de Encargos, item 4.1.3.8 (24h vias especiais, 48h demais), e Anexo 3 – "
        "Tabela 3-4 (IPA, peso 0,20 no ID).",
        "- Turno = prefixo da equipe; o IPA é atribuído à equipe que fechou a OS.",
        "- Erros de lançamento: regra de conferência entre serviço executado e aba Material, por OS (último fechamento) "
        "e inspeção, não um indicador do sistema.",
        f"- Projeção do mês assume o volume médio desde 01/{agora:%m}; vencidas em aberto contam como atrasos futuros.",
        f"- Aba gerada automaticamente pelo robô às {agora:%Hh%M}. O plano de ação é montado por regras a partir dos "
        "números; revise responsáveis e prazos antes de repassar.",
    ]))

    dados = {"os_fechadas": total, "no_prazo": no_prazo, "fora_do_prazo": fora, "ipa": ipa, "nota": nota(ipa),
             "dias": dias, "ultimo": ultimo, "extra": R}
    return dados, "\n\n".join(S) + "\n"


def materiais(dfm, mat):
    if mat.empty or "ID da Ordem/Inspeção" not in mat:
        return None
    reg = dfm[sem_obra_mask(dfm)].sort_values("fech").drop_duplicates(["Tipo", "ID"], keep="last").copy()
    m = mat.rename(columns={"ID da Ordem/Inspeção": "ID"}).merge(reg[["Tipo", "ID"]], on=["Tipo", "ID"])
    m["qtd"] = pd.to_numeric(m["Quantidade"], errors="coerce").fillna(0)
    m["tp"] = m["Tipo do Produto"].astype(str).str.upper().str.strip()
    m["prod"] = m["Produto"].astype(str).str.upper()

    def grupo(r):
        t, p = r["tp"], r["prod"]
        if t == "RELE":
            return "Relé de telegestão" if "TELEG" in p else "Relé fotoelétrico"
        if t == "CONECTOR":
            return "Conector perfurante" if "PERFUR" in p else "Conector"
        if t == "CABO":
            return "Cabo (PP, duplex, quadruplex)"
        if t in ("FITA ISOLANTE", "FITA"):
            return "Fita isolante" if "ISOL" in p or t == "FITA ISOLANTE" else "Fita"
        if t in ("LAMPADA", "LUMINARIA") or "PROJETOR" in p:
            return "Luminária / projetor / lâmpada LED"
        if t == "BASE":
            return "Base de relé"
        return t.capitalize()

    m["grupo"] = m.apply(grupo, axis=1)
    m["incl"] = m["Ação"].astype(str).str.startswith("Inclu")
    unid = m.groupby("grupo")["Unidade de Medida"].agg(lambda s: "m" if (s == "Metro").mean() > 0.5 else "un")
    gi = m[m["incl"]].groupby("grupo")["qtd"].sum()
    gr = m[~m["incl"]].groupby("grupo")["qtd"].sum()
    nreg = m.groupby("grupo").apply(lambda d: d[["Tipo", "ID"]].drop_duplicates().shape[0]).sort_values(ascending=False)

    def q(v, u):
        return "—" if not v else f"{br(v, 1 if v % 1 else 0)} {u}"

    grupos = [(g, q(gi.get(g, 0), unid[g]), q(gr.get(g, 0), unid[g]), br(int(n))) for g, n in nreg.head(10).items()]

    # erros por registro
    reg["svc"] = reg["Serviços Executados"].astype(str).str.upper()
    reg["exe"] = reg["Ocorrência"].astype(str).str.upper().eq("EXECUTADO")
    chave = list(zip(reg["Tipo"], reg["ID"]))
    por = m.groupby(["Tipo", "ID"])
    tem = set(por.groups)
    rele = m[m["tp"] == "RELE"]
    r_incl = set(map(tuple, rele[rele["incl"]][["Tipo", "ID"]].values.tolist()))
    r_ret = set(map(tuple, rele[~rele["incl"]][["Tipo", "ID"]].values.tolist()))
    incl_any = set(map(tuple, m[m["incl"]][["Tipo", "ID"]].values.tolist()))
    reg["has"] = [k in tem for k in chave]
    reg["rincl"] = [k in r_incl for k in chave]
    reg["rret"] = [k in r_ret for k in chave]
    reg["incl"] = [k in incl_any for k in chave]
    srv_rele = reg["svc"].str.contains(r"- RELÉ|- RELE\b", regex=True)
    troca = reg["svc"].str.contains("SUBSTITUI|INSTALAÇÃO|REPOSIÇÃO", regex=True)
    reg["troca"] = reg["exe"] & troca & ~reg["has"]
    reg["rele_srv"] = reg["exe"] & srv_rele & reg["has"] & ~(reg["rincl"] | reg["rret"])
    reg["rele_incl"] = reg["rincl"] & ~reg["rret"]
    reg["rele_sem_srv"] = (reg["rincl"] | reg["rret"]) & ~srv_rele
    reg["nexe_mat"] = ~reg["exe"] & reg["incl"]
    regras = ["troca", "rele_srv", "rele_incl", "rele_sem_srv", "nexe_mat"]
    reg["erro"] = reg[regras].any(axis=1)
    reg["equipe"] = reg["Equipe"].astype(str).str.strip()
    eq = reg.groupby("equipe").agg(reg=("erro", "size"), erro=("erro", "sum"), **{c: (c, "sum") for c in regras})
    eq["perc"] = 100 * eq["erro"] / eq["reg"]
    eq = eq[(eq["reg"] >= 80) & (eq["erro"] > 0)].sort_values("perc", ascending=False)
    tot = {c: int(reg[c].sum()) for c in regras + ["erro"]}
    tot["reg"] = len(reg)
    tot["perc"] = pct(tot["erro"], tot["reg"])

    os_exe = reg[(reg["Tipo"] == "Ordem de Serviço") & reg["exe"]]
    sem_mat = os_exe[~os_exe["has"]]
    n_rele = int(sem_mat["svc"].str.contains(r"- RELÉ|- RELE\b", regex=True).sum())
    outras = [f"{br(len(sem_mat))} OS executadas sem nenhum material, {br(n_rele)} delas com serviço de relé."]
    if "Inexistente" in m:
        inex = m[(~m["incl"]) & (m["Inexistente"].astype(str) == "Sim")]
        outras.append(f"{br(int((inex['Tipo'] == 'Ordem de Serviço').sum()))} retiradas de material marcado como "
                      f"inexistente no poste em OS (mais {br(int((inex['Tipo'] == 'Inspeção').sum()))} em inspeções).")
    fita = m[(m["grupo"] == "Fita isolante") & m["incl"] & (m["qtd"] >= 10)]
    outras.append(f"{br(len(fita))} lançamentos de fita isolante com 10 m ou mais.")
    return {"grupos": grupos, "equipes": eq, "tot": tot, "outras": outras, "regras": regras,
            "n_incl": int(m["incl"].sum()), "n_ret": int((~m["incl"]).sum()),
            "nome": {"troca": "Troca sem material", "rele_srv": "Relé no serviço sem relé lançado",
                     "rele_incl": "Relé incluído sem retirada", "rele_sem_srv": "Relé lançado sem serviço de relé",
                     "nexe_mat": "Não executada com material"},
            "curto": {"troca": "trocas sem material", "rele_srv": "relés no serviço sem relé lançado",
                      "rele_incl": "relés incluídos sem retirada", "rele_sem_srv": "relés lançados sem serviço de relé",
                      "nexe_mat": "não executadas com material"}}


def backlog(ab, agora):
    if ab is None or ab.empty or "Prioridade" not in ab:
        with open(Path(__file__).parent / "diagnostico.txt", "a", encoding="utf-8") as f:
            f.write(f"Aviso: OS em aberto sem coluna Prioridade. Colunas: {list(ab.columns) if ab is not None else None}\n")
        return None
    if "Tipo" in ab:
        ab = ab[ab["Tipo"].astype(str).str.contains("Ordem", case=False)]
    obra = ab[~sem_obra_mask(ab)]
    v = ab[sem_obra_mask(ab)].copy()
    agora_ts = pd.Timestamp(agora)
    v["esp"] = v["Prioridade"].map(eh_especial)
    v["venc"] = v["ab"] + pd.to_timedelta(v["esp"].map({True: 24, False: 48}), unit="h")
    v["vencida"] = v["venc"] < agora_ts
    v["v24"] = ~v["vencida"] & (v["venc"] < agora_ts + pd.Timedelta(hours=24))
    eqcol = "Equipe" if "Equipe" in v else None
    if eqcol:
        e = v[eqcol].fillna("").astype(str).str.strip()
        sem_eq = e.isin(["", "nan", "-", "None"])
        if "Status" in v:
            sem_eq |= v["Status"].fillna("").astype(str).str.strip().str.lower().eq("aberta")
        v["quem"] = e.where(~sem_eq, "Sem equipe (status Aberta)")
    else:
        v["quem"] = "Sem informação de equipe"
    g = v.groupby("quem").agg(n=("vencida", "size"), venc=("vencida", "sum"), v24=("v24", "sum"))
    sem = g.loc["Sem equipe (status Aberta)"] if "Sem equipe (status Aberta)" in g.index else None
    outros = g.drop(index="Sem equipe (status Aberta)", errors="ignore").sort_values(["venc", "n"], ascending=False)
    linhas = []
    if sem is not None:
        linhas.append(("Sem equipe (status Aberta)", br(int(sem.n)), br(int(sem.venc)), br(int(sem.v24))))
    top, resto = outros.head(4), outros.iloc[4:]
    linhas += [(k, br(int(r.n)), br(int(r.venc)), br(int(r.v24))) for k, r in top.iterrows()]
    if len(resto):
        linhas.append(("Outras equipes", br(int(resto.n.sum())), br(int(resto.venc.sum())), br(int(resto.v24.sum()))))
    linhas.append(("**Total**", f"**{br(len(v))}**", f"**{br(int(v['vencida'].sum()))}**", f"**{br(int(v['v24'].sum()))}**"))
    vse = v[(v["quem"] == "Sem equipe (status Aberta)") & v["vencida"]]
    idade = ""
    if len(vse):
        atraso = (agora_ts - vse["venc"]).dt.total_seconds() / 86400
        faixas = [(atraso <= 1).sum(), ((atraso > 1) & (atraso <= 3)).sum(), ((atraso > 3) & (atraso <= 7)).sum(), (atraso > 7).sum()]
        idade = (f"{br(int(faixas[0]))} venceram nas últimas 24h, {br(int(faixas[1]))} entre 1 e 3 dias, "
                 f"{br(int(faixas[2]))} entre 3 e 7 dias e {br(int(faixas[3]))} há mais de 7 dias")
    pior = (outros.index[0], int(outros.iloc[0].venc)) if len(outros) and outros.iloc[0].venc > 0 else None
    idade_obra = (agora_ts - obra["ab"]).dt.days if len(obra) else pd.Series(dtype=float)
    return {"abertas": len(v), "esp": int(v["esp"].sum()), "vencidas": int(v["vencida"].sum()),
            "venc_esp": int((v["vencida"] & v["esp"]).sum()),
            "venc_7d": int((v["vencida"] & ((agora_ts - v["venc"]) > pd.Timedelta(days=7))).sum()),
            "vence24": int(v["v24"].sum()),
            "sem_equipe": int(sem.n) if sem is not None else 0,
            "venc_sem_equipe": int(sem.venc) if sem is not None else 0,
            "esp_sem_equipe": int(v[v["quem"] == "Sem equipe (status Aberta)"]["esp"].sum()),
            "pior_eq": pior, "tabela": linhas, "idade": idade,
            "obras": len(obra), "obras90": int((idade_obra > 90).sum()), "obras180": int((idade_obra > 180).sum())}


# ---------------------------------------------------------------- gravação
def main(arquivo, arquivo_abertas=None):
    agora = datetime.now(TZ).replace(tzinfo=None)
    df, mat = ler_relatorio(arquivo)
    abertas = ler_abertas(arquivo_abertas)

    indice = json.loads((HIST / "indice.json").read_text(encoding="utf-8"))
    hoje = agora.strftime("%Y-%m-%d")
    autos_hoje = []
    for it in indice:
        if it["id"].startswith(hoje):
            try:
                if json.loads((HIST / f"{it['id']}.json").read_text(encoding="utf-8")).get("auto"):
                    autos_hoje.append(it["id"])
            except FileNotFoundError:
                autos_hoje.append(it["id"])
    restantes = [it for it in indice if it["id"] not in autos_hoje]
    ant = json.loads((HIST / f"{restantes[0]['id']}.json").read_text(encoding="utf-8")) if restantes else {}

    d, md = analisar(df, mat, abertas, agora, ant)
    mes = agora.strftime("%m/%Y")
    mesmo_mes = ant.get("mes") == mes
    novo_id = agora.strftime("%Y-%m-%d_%H%M")
    dados = {
        "titulo": ant.get("titulo", "Análise do Indicador de Prazo (Manaus Luz)"),
        "atualizado_em": agora.strftime("%d/%m/%Y %Hh%M"),
        "periodo": f"01/{agora:%m} a {agora:%d/%m/%Y}",
        "mes": mes,
        "os_fechadas": d["os_fechadas"],
        "no_prazo": d["no_prazo"],
        "fora_do_prazo": d["fora_do_prazo"],
        "ipa": d["ipa"],
        "nota": d["nota"],
        "ipa_anterior": ant.get("ipa") if mesmo_mes else None,
        "atualizado_anterior": ant.get("atualizado_em") if mesmo_mes else None,
        "doc": ant.get("doc", ""),
        "dias": d["dias"],
        "auto": True,
        "extra": d["extra"],
    }
    assert sum(x["os"] for x in d["dias"]) == d["os_fechadas"]
    assert sum(x["atrasos"] for x in d["dias"]) == d["fora_do_prazo"]

    for antigo in autos_hoje:
        for ext in (".json", ".md"):
            (HIST / f"{antigo}{ext}").unlink(missing_ok=True)
    (HIST / f"{novo_id}.json").write_text(json.dumps(dados, ensure_ascii=False, indent=1), encoding="utf-8")
    (HIST / f"{novo_id}.md").write_text(md, encoding="utf-8")
    restantes.insert(0, {"id": novo_id, "rotulo": f"Até {agora:%d/%m %H:%M}"})
    (HIST / "indice.json").write_text(json.dumps(restantes, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"IPA {br(d['ipa'], 2)}% ({d['no_prazo']}/{d['os_fechadas']}), aba {novo_id}"
          + (f", substituiu {', '.join(autos_hoje)}" if autos_hoje else "")
          + ("" if abertas is not None else " (sem backlog)"))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "relatorio.xlsx",
         sys.argv[2] if len(sys.argv) > 2 else None)
