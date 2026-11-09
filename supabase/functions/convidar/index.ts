// @ts-nocheck
// FAROL IPA · Edge Function "convidar" (mesmo sistema de senhas do Provius)
// Cria o acesso de uma pessoa com uma SENHA PROVISÓRIA (vale 24 horas) e devolve essa senha
// para o administrador repassar. No primeiro acesso a pessoa é obrigada a criar a própria senha.
// Também gera uma nova senha provisória quando alguém esquece a senha.
// Não depende de email. Só o administrador consegue usar.
// Como publicar: Supabase (projeto do Farol) > Edge Functions > Deploy a new function > Via Editor >
// nome "convidar" > apague o exemplo, cole este arquivo inteiro > Deploy.
// (Se a função já existir: abra "convidar" > Code > cole por cima > Deploy.)
import { createClient } from 'npm:@supabase/supabase-js@2';

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};
const resp = (corpo: unknown, status = 200) =>
  new Response(JSON.stringify(corpo), { status, headers: { ...CORS, 'Content-Type': 'application/json' } });

const VALIDADE_H = 24;

// senha provisória fácil de ditar: letras e números sem os que confundem (0/O, 1/l/I)
function senhaProvisoria() {
  const L = 'abcdefghjkmnpqrstuvwxyz', N = '23456789', T = L + N + L.toUpperCase();
  const r = (s: string) => s[crypto.getRandomValues(new Uint32Array(1))[0] % s.length];
  let s = ''; for (let i = 0; i < 6; i++) s += r(T);
  return 'Fa' + r(N) + r(N) + '-' + s;   // ex.: Fa47-kR9mTz
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: CORS });
  try {
    const url = Deno.env.get('SUPABASE_URL')!;
    const anon = Deno.env.get('SUPABASE_ANON_KEY')!;
    const servico = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;

    const auth = req.headers.get('Authorization') || '';
    const comoUsuario = createClient(url, anon, { global: { headers: { Authorization: auth } } });
    const { data: u } = await comoUsuario.auth.getUser();
    if (!u?.user) return resp({ erro: 'Entre no Farol de novo.' }, 401);

    const adm = createClient(url, servico, { auth: { persistSession: false } });
    const { data: eu } = await adm.from('perfis').select('id,admin,ativo').eq('id', u.user.id).maybeSingle();
    if (!eu?.admin || !eu?.ativo) return resp({ erro: 'Só o administrador cria acessos.' }, 403);

    const b = await req.json();
    const email = String(b.email || '').trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return resp({ erro: 'Email inválido.' }, 400);
    const expira = new Date(Date.now() + VALIDADE_H * 3600e3).toISOString();
    const senha = senhaProvisoria();
    const { data: existente } = await adm.from('perfis').select('id').eq('email', email).maybeSingle();

    if (b.acao === 'nova_senha') {
      if (!existente) return resp({ erro: 'Pessoa não encontrada.' }, 404);
      if (existente.id === eu.id) return resp({ erro: 'Para trocar a sua própria senha, use "Trocar senha".' }, 400);
      const { error } = await adm.auth.admin.updateUserById(existente.id, { password: senha });
      if (error) return resp({ erro: error.message }, 400);
      await adm.from('perfis').update({ senha_ok: false, convite_enviado: new Date().toISOString(), convite_expira: expira }).eq('id', existente.id);
      return resp({ ok: true, email, senha, expira });
    }

    // criar
    const nome = String(b.nome || '').trim();
    if (!nome) return resp({ erro: 'Informe o nome.' }, 400);
    if (existente) return resp({ erro: 'Já existe uma pessoa com esse email.' }, 400);
    const { data: c, error } = await adm.auth.admin.createUser({ email, password: senha, email_confirm: true, user_metadata: { nome } });
    if (error) return resp({ erro: /already/i.test(error.message) ? 'Esse email já tem conta no Supabase.' : error.message }, 400);
    const { error: e2 } = await adm.from('perfis').insert({ id: c.user.id, nome, email, admin: !!b.admin, senha_ok: false,
      convite_enviado: new Date().toISOString(), convite_expira: expira });
    if (e2) { await adm.auth.admin.deleteUser(c.user.id); return resp({ erro: e2.message }, 400); }
    return resp({ ok: true, email, senha, expira });
  } catch (e) {
    return resp({ erro: String(e?.message || e) }, 500);
  }
});
