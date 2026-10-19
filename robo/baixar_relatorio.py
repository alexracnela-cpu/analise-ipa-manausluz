"""Entra no GisWorks (Manaus Luz) e baixa o relatório
"Ordens Executadas e Inspeções/Materiais por Município" do mês corrente.

Também baixa (se conseguir) "Todas as Ordens em Aberto ou em Andamento
atualmente por Município", usado no backlog da análise.

Uso: python robo/baixar_relatorio.py saida.xlsx [abertas.xlsx]
Precisa das variáveis de ambiente GISWORKS_USUARIO e GISWORKS_SENHA
(no GitHub ficam em Settings > Secrets and variables > Actions).
"""
import base64
import os
import sys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

from playwright.sync_api import sync_playwright

BASE = os.environ.get("GISWORKS_URL", "https://manaus.gestoriluminacaopublica.com.br")
TZ = ZoneInfo("America/Manaus")

JS_GERAR = r"""
async ({ini, fim}) => {
  const dorme = ms => new Promise(r => setTimeout(r, ms));
  // 1) abre o modal do relatório
  const links = [...document.querySelectorAll('a')];
  const link = links.find(a => /OrdensInspecao\s*\(/.test(a.getAttribute('onclick') || ''))
            || links.find(a => /Ordens Executadas\s+e\s+Inspe/i.test(a.textContent || ''));
  if (!link) throw new Error('Link do relatório "Ordens Executadas e Inspeções" não encontrado em /Relatorio');
  link.click();
  let form = null;
  for (let i = 0; i < 100 && !form; i++) { await dorme(200); form = document.querySelector('.modal.in form, .modal.show form'); }
  if (!form) throw new Error('O formulário do relatório não abriu');

  // 2) município Manaus
  const sel = form.querySelector('select[name=MunicipioId]');
  let municipio = null;
  if (sel) {
    const opt = [...sel.options].find(o => /manaus/i.test(o.text));
    if (!opt) throw new Error('Município Manaus não encontrado no formulário');
    sel.value = opt.value; municipio = opt.value;
    sel.dispatchEvent(new Event('change', {bubbles: true}));
    await dorme(1500);
  }

  // 3) monta os campos: fechamento e inspeção no período, sem filtro de criação
  const fd = new FormData(form);
  for (const k of [...fd.keys()]) if (/Criacao/i.test(k)) fd.delete(k);
  if (municipio) fd.set('MunicipioId', municipio);
  const nomes = [...form.elements].map(e => e.name).filter(Boolean);
  for (const n of new Set(nomes)) {
    if (/^DataInicial(Fechamento|Inspecao)$/i.test(n)) fd.set(n, ini);
    if (/^DataFinal(Fechamento|Inspecao)$/i.test(n)) fd.set(n, fim);
    if (/^HoraInicial(Fechamento|Inspecao)$/i.test(n)) fd.set(n, '00:00:00');
    if (/^HoraFinal(Fechamento|Inspecao)$/i.test(n)) fd.set(n, '23:59:59');
  }
  const url = form.getAttribute('action') || '/Relatorio/OrdensInspecao';

  // 4) gera e devolve o xlsx em base64
  const resp = await fetch(url, {method: 'POST', body: fd, credentials: 'include'});
  if (!resp.ok) throw new Error('GisWorks respondeu ' + resp.status + ' ao gerar o relatório');
  const blob = await resp.blob();
  const b64 = await new Promise((ok, erro) => {
    const r = new FileReader();
    r.onload = () => ok(String(r.result).split(',')[1]);
    r.onerror = erro;
    r.readAsDataURL(blob);
  });
  return {tipo: resp.headers.get('content-type'), tamanho: blob.size, b64};
}
"""

JS_ABERTAS = r"""
async () => {
  const dorme = ms => new Promise(r => setTimeout(r, ms));
  const links = [...document.querySelectorAll('a')];
  const link = links.find(a => /TodasOrdensAbertas\s*\(/.test(a.getAttribute('onclick') || ''))
            || links.find(a => /todas as\s+Ordens\s+em Aberto ou em Andamento\s+atualmente/i.test(a.textContent || ''));
  if (!link) throw new Error('Link do relatório de OS em aberto não encontrado');
  link.click();
  let form = null;
  for (let i = 0; i < 100 && !form; i++) { await dorme(200); form = document.querySelector('.modal.in form, .modal.show form'); }
  if (!form) throw new Error('O formulário de OS em aberto não abriu');
  const sel = form.querySelector('select[name=MunicipioId]');
  if (sel) {
    const opt = [...sel.options].find(o => /manaus/i.test(o.text));
    if (opt) { sel.value = opt.value; sel.dispatchEvent(new Event('change', {bubbles: true})); }
    await dorme(1000);
  }
  const fd = new FormData(form);
  if (sel) fd.set('MunicipioId', sel.value);
  fd.set('Relatorio', 'TodasOrdensAbertas');
  const resp = await fetch('/Relatorio/GerarRelatorio', {method: 'POST', body: new URLSearchParams(fd), credentials: 'include'});
  if (!resp.ok) throw new Error('GisWorks respondeu ' + resp.status + ' ao gerar OS em aberto');
  const blob = await resp.blob();
  const b64 = await new Promise((ok, erro) => {
    const r = new FileReader();
    r.onload = () => ok(String(r.result).split(',')[1]);
    r.onerror = erro;
    r.readAsDataURL(blob);
  });
  return {tipo: resp.headers.get('content-type'), tamanho: blob.size, b64};
}
"""


def main(saida: str, saida_abertas: str | None = None) -> None:
    usuario = os.environ["GISWORKS_USUARIO"]
    senha = os.environ["GISWORKS_SENHA"]
    agora = datetime.now(TZ)
    ini = agora.replace(day=1).strftime("%Y-%m-%d")
    # o GisWorks pode tratar a data final como exclusiva: pede até amanhã
    # (fechamentos futuros não existem) e o recorte exato é feito no Python
    fim = (agora + timedelta(days=1)).strftime("%Y-%m-%d")

    with sync_playwright() as p:
        nav = p.chromium.launch()
        pagina = nav.new_page(locale="pt-BR", timezone_id="America/Manaus")
        pagina.set_default_timeout(60_000)
        try:
            pagina.goto(BASE + "/")
            pagina.locator("input[type=password]").wait_for()
            campo_usuario = pagina.locator(
                "form input:not([type=password]):not([type=hidden]):not([type=checkbox]):not([type=submit])"
            ).first
            campo_usuario.fill(usuario)
            pagina.locator("input[type=password]").fill(senha)
            botao = pagina.locator("button:has-text('Login'), input[type=submit]").first
            botao.click()
            try:
                pagina.locator("input[type=password]").wait_for(state="detached", timeout=30_000)
            except Exception:
                raise RuntimeError("Login no GisWorks falhou (usuário/senha errados ou tela mudou)")

            pagina.goto(BASE + "/Relatorio")
            pagina.wait_for_load_state("networkidle")
            r = pagina.evaluate(JS_GERAR, {"ini": ini, "fim": fim})
            ra = None
            if saida_abertas:
                try:
                    pagina.goto(BASE + "/Relatorio")
                    pagina.wait_for_load_state("networkidle")
                    ra = pagina.evaluate(JS_ABERTAS)
                except Exception as e:  # o backlog é opcional: a análise sai sem ele
                    print(f"Aviso: não consegui baixar as OS em aberto ({e})")
        except Exception:
            pagina.screenshot(path="erro_gisworks.png", full_page=True)
            raise
        finally:
            nav.close()

    dados = base64.b64decode(r["b64"])
    if not dados.startswith(b"PK"):
        raise RuntimeError(f"O GisWorks não devolveu um Excel (tipo {r['tipo']}, {r['tamanho']} bytes)")
    with open(saida, "wb") as f:
        f.write(dados)
    print(f"Relatório baixado: {saida} ({len(dados):,} bytes), fechamentos {ini} a {fim}")
    if ra:
        da = base64.b64decode(ra["b64"])
        if da.startswith(b"PK"):
            with open(saida_abertas, "wb") as f:
                f.write(da)
            print(f"OS em aberto baixadas: {saida_abertas} ({len(da):,} bytes)")
        else:
            print(f"Aviso: OS em aberto não vieram em Excel (tipo {ra['tipo']})")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "relatorio.xlsx",
         sys.argv[2] if len(sys.argv) > 2 else None)
