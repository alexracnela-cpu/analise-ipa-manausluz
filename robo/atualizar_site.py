"""Calcula o IPA a partir do relatório do GisWorks e cria a aba nova do site.

Uso: python robo/atualizar_site.py relatorio.xlsx

Regras (as mesmas das análises manuais):
- base = OS ("Ordem de Serviço") fechadas no mês, exceto origem Obra / reclamante "Obra"
  (OS de vias fechadas por equipes de obra continuam na base);
- inspeções fora do IPA;
- no prazo = campo "Fechada no Prazo" = "Sim";
- IPA com 2 casas.

O robô mantém UMA aba automática por dia: a cada execução substitui a aba
automática do mesmo dia. Abas feitas à mão (sem "auto": true) nunca são tocadas.
"""
import json
import math
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
HIST = RAIZ / "historico"
TZ = ZoneInfo("America/Manaus")

# faixas de nota do IPA (Anexo 3). Ajuste aqui se a tabela do contrato tiver mais faixas.
FAIXAS_NOTA = [(95.0, "1,0"), (90.0, "0,8"), (0.0, "0,0")]
ORIGENS_RECLAMACAO = {"Sistema - Criada no Call Center", "App - Android", "App - iOS", "Web"}


def br(n, casas=0):
    s = f"{n:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(sim, total):
    return 100.0 * sim / total if total else 0.0


def nota(ipa):
    for limite, n in FAIXAS_NOTA:
        if round(ipa, 2) >= limite:
            return n
    return FAIXAS_NOTA[-1][1]


def ler(arquivo):
    df = pd.read_excel(arquivo, sheet_name="Relatório", header=1)
    df.columns = [str(c).strip() for c in df.columns]
    df["fech"] = pd.to_datetime(df["Data Fechamento"], dayfirst=True, errors="coerce")
    return df


def linha_recorte(nome, d):
    s = int((d["Fechada no Prazo"] == "Sim").sum())
    p = pct(s, len(d))
    return f"| {nome} | {br(len(d))} | {br(p, 2)}% | {nota(p)} |"


def resumo_grupo(d, col, ordem=None):
    g = d.groupby(col).agg(os=("prazo", "size"), sim=("prazo", "sum"))
    g["atrasos"] = g["os"] - g["sim"]
    g["ipa"] = 100 * g["sim"] / g["os"]
    return g.sort_values("atrasos", ascending=False) if ordem is None else g.reindex(ordem).dropna()


def main(arquivo):
    agora = datetime.now(TZ)
    df = ler(arquivo)
    ini_mes = pd.Timestamp(agora.year, agora.month, 1)
    limite = pd.Timestamp(agora.replace(tzinfo=None))
    df = df[(df["fech"] >= ini_mes) & (df["fech"] <= limite)]

    os_ = df[df["Tipo"] == "Ordem de Serviço"]
    obra = (os_["Origem"].astype(str).str.strip() == "Obra") | (
        os_["Reclamante"].astype(str).str.strip().str.upper() == "OBRA")
    base = os_[~obra].copy()
    if base.empty:
        raise SystemExit("Relatório sem OS no período — nada a publicar.")
    base["prazo"] = base["Fechada no Prazo"] == "Sim"

    total = len(base)
    no_prazo = int(base["prazo"].sum())
    fora = total - no_prazo
    ipa = round(pct(no_prazo, total), 2)
    ultimo = base["fech"].max()
    n_insp = int((df["Tipo"] == "Inspeção").sum())
    n_obra = int(obra.sum())
    obra_eq = int(base["Equipe"].astype(str).str.upper().str.contains("OBRA").sum())

    # gráfico: OS e atrasos por dia de fechamento
    dias = []
    por_dia = base.groupby(base["fech"].dt.day)
    for dia in range(1, ultimo.day + 1):
        if dia in por_dia.groups:
            g = por_dia.get_group(dia)
            dias.append({"dia": f"{dia:02d}", "os": len(g), "atrasos": int((~g["prazo"]).sum())})
        else:
            dias.append({"dia": f"{dia:02d}", "os": 0, "atrasos": 0})

    # aba anterior (a mais recente que não seja a automática de hoje)
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
    mes = agora.strftime("%m/%Y")
    mesmo_mes = ant.get("mes") == mes

    novo_id = agora.strftime("%Y-%m-%d_%H%M")
    dados = {
        "titulo": ant.get("titulo", "Análise do Indicador de Prazo (Manaus Luz)"),
        "atualizado_em": agora.strftime("%d/%m/%Y %Hh%M"),
        "periodo": f"01/{agora:%m} a {agora:%d/%m/%Y}",
        "mes": mes,
        "os_fechadas": total,
        "no_prazo": no_prazo,
        "fora_do_prazo": fora,
        "ipa": ipa,
        "nota": nota(ipa),
        "ipa_anterior": ant.get("ipa") if mesmo_mes else None,
        "atualizado_anterior": ant.get("atualizado_em") if mesmo_mes else None,
        "doc": ant.get("doc", ""),
        "dias": dias,
        "auto": True,
    }
    assert sum(d["os"] for d in dias) == total and sum(d["atrasos"] for d in dias) == fora

    # ---------- texto da aba ----------
    p95 = max(0, math.ceil(0.95 * total - no_prazo - 1e-9))
    saldo90 = math.floor(total * 0.10 - fora + 1e-9)
    hoje_d = base[base["fech"].dt.date == agora.date()]
    linhas = []
    linhas.append(
        f"**Período considerado:** fechamentos de 01/{agora:%m/%Y} 00h00 até {ultimo:%d/%m/%Y} às "
        f"{ultimo:%Hh%M} (último fechamento no relatório, gerado automaticamente às {agora:%Hh%M}). "
        f"Base: todas as OS fechadas no período, exceto as de origem Obra ({br(n_obra)} OS, reclamante \"Obra\")"
        + (f"; as {br(obra_eq)} OS de vias fechadas por equipes de obra estão incluídas" if obra_eq else "")
        + f". As inspeções ({br(n_insp)}) não entram no IPA. O total confere: "
          f"{br(no_prazo)} + {br(fora)} = {br(total)}.")
    linhas.append("<!-- GRAFICO -->")
    linhas.append("## Resumo executivo")
    var = ""
    if dados["ipa_anterior"] is not None:
        dv = ipa - dados["ipa_anterior"]
        if abs(dv) < 0.005:
            var = f" Em relação à aba anterior ({dados['atualizado_anterior']}), o índice ficou estável."
        else:
            var = (f" Em relação à aba anterior ({dados['atualizado_anterior']}), o índice "
                   f"{'subiu' if dv > 0 else 'caiu'} {'+' if dv > 0 else '−'}{br(abs(dv), 2)} p.p. "
                   f"(de {br(dados['ipa_anterior'], 2)}%).")
    linhas.append(
        f"O IPA de {agora:%m/%Y} está em **{br(ipa, 2)}% — nota {nota(ipa)}** "
        f"({br(no_prazo)} de {br(total)} OS no prazo, {br(fora)} fora do prazo).{var}"
        + (f" Hoje, até {ultimo:%Hh%M}, foram {br(len(hoje_d))} fechamentos com "
           f"{br(int((~hoje_d['prazo']).sum()))} atrasos." if len(hoje_d) else ""))
    if ipa >= 95:
        meta = f"≥95,00% (nota 1,0) atingido hoje, com folga de {br(math.floor(total*0.05 - fora + 1e-9))} atrasos."
    else:
        meta = (f"Para estar em 95,00% hoje seriam necessários {br(p95)} atrasos a menos. "
                + (f"≥90,00% (nota 0,8) tem folga de {br(saldo90)} atrasos no total atual."
                   if saldo90 >= 0 else f"Faltam {br(-saldo90)} atrasos a menos para voltar a 90,00%."))
    linhas.append(f"**Meta:** {meta}")

    linhas.append("## IPA com recortes")
    rec = base[base["Origem"].isin(ORIGENS_RECLAMACAO)]
    unica = base.sort_values("fech").drop_duplicates("ID", keep="last")
    execu = base[base["Ocorrência"].astype(str).str.upper() == "EXECUTADO"]
    linhas.append("\n".join([
        f"| Recorte (01/{agora:%m}–{agora:%d/%m}) | OS | No prazo | Nota IPA |",
        "| --- | --- | --- | --- |",
        linha_recorte("Todas as OS, exceto origem Obra (base adotada)", base),
        linha_recorte("Cada OS uma vez, pelo último fechamento", unica),
        linha_recorte("Só reclamações de usuários (Call Center, App, Web)", rec),
        linha_recorte("Só OS executadas", execu),
    ]))

    def item(nome, g):
        return f"{nome} {br(g.ipa, 2)}% ({br(g.atrasos)} atrasos em {br(g.os)})"

    pri = resumo_grupo(base, "Prioridade")
    bullets = []
    if len(pri):
        bullets.append("- **Prioridade:** " + "; ".join(
            item("vias especiais" if "especiais" in str(k).lower() and "não" not in str(k).lower()
                 else "demais vias" if "não" in str(k).lower() else str(k).strip(), r)
            for k, r in pri.iterrows()) + ".")
    ori = resumo_grupo(base, "Origem").sort_values("ipa", ascending=False)
    bullets.append("- **Origem:** " + "; ".join(item(str(k), r) for k, r in ori.iterrows()) + ".")
    exe = resumo_grupo(base.assign(oc=base["Ocorrência"].astype(str).str.capitalize()), "oc")
    bullets.append("- **Executada x não executada:** " + "; ".join(item(str(k).lower().replace("executado", "executadas"), r) for k, r in exe.iterrows()) + ".")
    zon = resumo_grupo(base[base["Área"].astype(str).str.upper().str.startswith("ZONA")], "Área")
    bullets.append("- **Zonas:** " + "; ".join(item(str(k).title(), r) for k, r in zon.iterrows()) + ".")
    linhas.append("\n".join(bullets))

    linhas.append("## Equipes com mais atrasos")
    eq = resumo_grupo(base, "Equipe")
    eq = eq[eq["atrasos"] > 0].head(10)
    tab = ["| Equipe | OS | No prazo | Atrasos |", "| --- | --- | --- | --- |"]
    tab += [f"| {str(k).strip()} | {br(r.os)} | {br(r.ipa, 2)}% | {br(r.atrasos)} |" for k, r in eq.iterrows()]
    linhas.append("\n".join(tab))
    linhas.append(f"*Aba gerada automaticamente pelo robô às {agora:%Hh%M} a partir do GisWorks. "
                  "A análise completa (causa raiz, materiais e plano de ação) fica no doc do IPA.*")
    md = "\n\n".join(linhas) + "\n"

    # ---------- grava ----------
    for antigo in autos_hoje:
        for ext in (".json", ".md"):
            (HIST / f"{antigo}{ext}").unlink(missing_ok=True)
    (HIST / f"{novo_id}.json").write_text(json.dumps(dados, ensure_ascii=False, indent=1), encoding="utf-8")
    (HIST / f"{novo_id}.md").write_text(md, encoding="utf-8")
    restantes.insert(0, {"id": novo_id, "rotulo": f"Até {agora:%d/%m %H:%M}"})
    (HIST / "indice.json").write_text(json.dumps(restantes, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"IPA {br(ipa, 2)}% ({no_prazo}/{total}), aba {novo_id}"
          + (f", substituiu {', '.join(autos_hoje)}" if autos_hoje else ""))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "relatorio.xlsx")
