"""Leitura e gravação das abas do Farol no Supabase (tabela "abas").

Só entra em ação quando os secrets SUPABASE_URL e SUPABASE_SECRET existem no GitHub.
Sem eles, o robô continua no modo antigo (arquivos em historico/ no repositório).

- baixar(): antes da análise, traz do banco o índice e os números das abas para
  historico/ (o robô usa a aba anterior para as comparações).
  Se o banco ainda estiver vazio, sobe primeiro todo o histórico que existe em historico/.
- gravar(): depois da análise, envia a aba nova e apaga a automática do mesmo dia que ela substituiu.
"""
import json
import os
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

URL = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
CHAVE = os.environ.get("SUPABASE_SECRET", "").strip()


def configurado():
    return bool(URL and CHAVE)


def _req(metodo, caminho, corpo=None, prefer=None):
    cab = {"apikey": CHAVE, "Content-Type": "application/json", "Accept": "application/json"}
    if CHAVE.startswith("eyJ"):  # chave service_role antiga (JWT)
        cab["Authorization"] = f"Bearer {CHAVE}"
    if prefer:
        cab["Prefer"] = prefer
    dados = json.dumps(corpo, ensure_ascii=False).encode("utf-8") if corpo is not None else None
    req = urllib.request.Request(f"{URL}/rest/v1/{caminho}", data=dados, method=metodo, headers=cab)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            txt = r.read().decode("utf-8")
            return json.loads(txt) if txt else None
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Supabase {metodo} {caminho}: {e.code} {e.read().decode('utf-8', 'replace')[:400]}") from None


def _upsert(linhas):
    for i in range(0, len(linhas), 20):
        _req("POST", "abas?on_conflict=id", linhas[i:i + 20], prefer="resolution=merge-duplicates,return=minimal")


def _subir_historico_local(hist: Path):
    """Primeira vez: copia para o banco tudo o que está em historico/."""
    arq = hist / "indice.json"
    if not arq.exists():
        return 0
    linhas = []
    for it in json.loads(arq.read_text(encoding="utf-8")):
        j, m = hist / f"{it['id']}.json", hist / f"{it['id']}.md"
        if not j.exists():
            continue
        linhas.append({"id": it["id"], "rotulo": it["rotulo"],
                       "dados": json.loads(j.read_text(encoding="utf-8")),
                       "md": m.read_text(encoding="utf-8") if m.exists() else ""})
    _upsert(linhas)
    return len(linhas)


def baixar(hist: Path):
    hist.mkdir(parents=True, exist_ok=True)
    abas = _req("GET", "abas?select=id,rotulo,dados&order=id.desc")
    if not abas:
        n = _subir_historico_local(hist)
        print(f"Supabase vazio: enviei {n} abas do histórico do repositório.")
        abas = _req("GET", "abas?select=id,rotulo,dados&order=id.desc") or []
    for a in abas:
        (hist / f"{a['id']}.json").write_text(json.dumps(a["dados"], ensure_ascii=False, indent=1), encoding="utf-8")
    (hist / "indice.json").write_text(
        json.dumps([{"id": a["id"], "rotulo": a["rotulo"]} for a in abas], ensure_ascii=False, indent=1),
        encoding="utf-8")


def gravar(novo_id, rotulo, dados, md, apagar=()):
    _upsert([{"id": novo_id, "rotulo": rotulo, "dados": dados, "md": md}])
    velhos = [x for x in apagar if x != novo_id]
    if velhos:
        lista = ",".join(urllib.parse.quote(f'"{x}"') for x in velhos)
        _req("DELETE", f"abas?id=in.({lista})", prefer="return=minimal")
