// paut-api — PAUT 앱용 Microsoft 로그인 중계 서버 (Cloudflare Worker) · v2
//
// 왜: iOS 홈 화면 앱(SPA)은 refresh token 이 24시간이면 끝나고, 숨은 iframe 갱신도 막혀 매일 로그인이 풀린다.
// 이 서버는 Azure 'Web' 플랫폼(client secret) 으로 로그인을 받아 refresh token 을 서버(KV)에 보관한다.
// 휴대폰에는 이 기기 전용 열쇠(device token)만 두고, Graph access token 은 필요할 때 여기서 받아 간다.
//
// 경로
//   GET  /                 상태·등록할 리디렉션 URI 안내
//   GET  /auth/login       Microsoft 로그인으로 보낸다 (?state=앱이 만든 값, ?hint=이메일)
//   GET  /auth/callback    Microsoft 가 돌아오는 곳. 토큰 교환 → 기기 열쇠 발급 → 앱으로 돌려보냄(#paut_device=…)
//   POST /auth/logout      이 기기 열쇠 폐기
//   GET  /api/me           { email }
//   POST /api/token        { access_token, expires_at, email }  (?force=1 이면 캐시 무시)
//
// 환경
//   vars:    AZURE_CLIENT_ID, APP_URL(https://…/paut-app/), APP_ORIGIN(https://….github.io)
//   secret:  AZURE_CLIENT_SECRET
//   KV:      STORE

const AUTHORITY = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const SCOPE = "offline_access User.Read Files.ReadWrite";
const LOGIN_TTL = 600; // 로그인 왕복 대기 10분

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (request.method === "OPTIONS") return cors(env, new Response(null, { status: 204 }));
    try {
      switch (`${request.method} ${url.pathname}`) {
        case "GET /":
          return home(url, env);
        case "GET /auth/login":
          return await login(url, env);
        case "GET /auth/callback":
          return await callback(url, env);
        case "POST /auth/logout":
          return cors(env, await logout(request, env));
        case "GET /api/me":
          return cors(env, await me(request, env));
        case "POST /api/token":
          return cors(env, await token(request, url, env));
        default:
          return cors(env, json({ error: "not_found" }, 404));
      }
    } catch (error) {
      console.error(error);
      return cors(env, json({ error: "server_error", message: String(error?.message ?? error) }, 500));
    }
  },
};

// ───────── 로그인 ─────────

function redirectUri(url) {
  return `${url.origin}/auth/callback`;
}

function home(url, env) {
  const ok = !!(env.AZURE_CLIENT_ID && env.AZURE_CLIENT_SECRET && env.STORE);
  return new Response(
    [
      "paut-api 동작 중",
      `설정 완료: ${ok ? "예" : "아니오 (AZURE_CLIENT_ID / AZURE_CLIENT_SECRET / KV 확인)"}`,
      "",
      "Azure 앱 등록 → 인증 → 플랫폼 추가 → 'Web' 에 아래 리디렉션 URI 를 넣으세요:",
      redirectUri(url),
      "",
      "앱 빌드 변수 PAUT_API_URL 에 넣을 주소:",
      url.origin,
    ].join("\n"),
    { headers: { "content-type": "text/plain; charset=utf-8" } },
  );
}

async function login(url, env) {
  const appState = (url.searchParams.get("state") ?? "").slice(0, 200);
  const hint = (url.searchParams.get("hint") ?? "").slice(0, 200);
  const state = randomToken(24);
  const verifier = randomToken(48);
  await env.STORE.put(`login:${state}`, JSON.stringify({ verifier, appState }), { expirationTtl: LOGIN_TTL });
  const q = new URLSearchParams({
    client_id: env.AZURE_CLIENT_ID,
    response_type: "code",
    redirect_uri: redirectUri(url),
    response_mode: "query",
    scope: SCOPE,
    state,
    code_challenge: await sha256b64url(verifier),
    code_challenge_method: "S256",
  });
  if (hint) q.set("login_hint", hint);
  return Response.redirect(`${AUTHORITY}/authorize?${q}`, 302);
}

async function callback(url, env) {
  const back = (params) => Response.redirect(`${env.APP_URL}#${new URLSearchParams(params)}`, 302);
  const err = url.searchParams.get("error");
  const state = url.searchParams.get("state") ?? "";
  const saved = state ? await env.STORE.get(`login:${state}`, "json") : null;
  if (state) await env.STORE.delete(`login:${state}`);
  const appState = saved?.appState ?? "";
  if (err) return back({ paut_error: `${err}: ${url.searchParams.get("error_description") ?? ""}`.slice(0, 300), paut_state: appState });
  if (!saved) return back({ paut_error: "로그인 시간이 지났거나 상태값이 없습니다. 다시 로그인하세요.", paut_state: appState });

  const tok = await tokenRequest(env, {
    grant_type: "authorization_code",
    code: url.searchParams.get("code") ?? "",
    redirect_uri: redirectUri(url),
    code_verifier: saved.verifier,
  });
  if (!tok.ok) return back({ paut_error: `토큰 교환 실패: ${tok.error}`, paut_state: appState });

  const profile = await graphMe(tok.data.access_token);
  if (!profile) return back({ paut_error: "Microsoft 계정 정보를 읽지 못했습니다.", paut_state: appState });

  // 한 사람용 서버: 처음 로그인한 계정이 주인. 다른 계정은 거부.
  const owner = await env.STORE.get("owner", "json");
  if (owner && owner.id !== profile.id) {
    return back({ paut_error: `이 서버는 ${owner.email} 계정 전용입니다.`, paut_state: appState });
  }
  if (!owner) await env.STORE.put("owner", JSON.stringify({ id: profile.id, email: profile.email }));

  await saveTokens(env, tok.data);
  const device = randomToken(32);
  await env.STORE.put(`dev:${await sha256b64url(device)}`, JSON.stringify({ email: profile.email, created: Date.now() }));
  return back({ paut_device: device, paut_email: profile.email, paut_state: appState });
}

async function logout(request, env) {
  const key = await deviceKey(request);
  if (key) await env.STORE.delete(key);
  return json({ ok: true });
}

// ───────── API ─────────

async function me(request, env) {
  const dev = await device(request, env);
  if (!dev) return json({ error: "device_unknown" }, 401);
  return json({ email: dev.email });
}

async function token(request, url, env) {
  const dev = await device(request, env);
  if (!dev) return json({ error: "device_unknown" }, 401);
  const force = url.searchParams.get("force") === "1";
  if (!force) {
    const cached = await env.STORE.get("at", "json");
    if (cached && cached.expires_at - Date.now() > 5 * 60_000) return json({ ...cached, email: dev.email });
  }
  const rt = await env.STORE.get("rt");
  if (!rt) return json({ error: "login_required", message: "서버에 Microsoft 로그인이 없습니다." }, 401);
  const tok = await tokenRequest(env, { grant_type: "refresh_token", refresh_token: rt });
  if (!tok.ok) {
    // invalid_grant = refresh token 이 정말 끝남(비밀번호 변경·권한 철회·오래 안 씀). 그 밖은 일시 오류.
    if (tok.code === "invalid_grant" || tok.code === "interaction_required") {
      return json({ error: "login_required", message: tok.error }, 401);
    }
    return json({ error: "upstream", message: tok.error }, 502);
  }
  const at = await saveTokens(env, tok.data);
  return json({ ...at, email: dev.email });
}

// ───────── 도우미 ─────────

async function tokenRequest(env, params) {
  const body = new URLSearchParams({
    client_id: env.AZURE_CLIENT_ID,
    client_secret: env.AZURE_CLIENT_SECRET,
    scope: SCOPE,
    ...params,
  });
  const res = await fetch(`${AUTHORITY}/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.access_token) {
    return { ok: false, code: data.error ?? `http_${res.status}`, error: `${data.error ?? res.status} ${data.error_description ?? ""}`.trim().slice(0, 300) };
  }
  return { ok: true, data };
}

/** refresh token(회전된 새 값) 과 access token 캐시를 저장한다. */
async function saveTokens(env, data) {
  if (data.refresh_token) await env.STORE.put("rt", data.refresh_token);
  const at = { access_token: data.access_token, expires_at: Date.now() + (Number(data.expires_in) || 3600) * 1000 };
  await env.STORE.put("at", JSON.stringify(at), { expirationTtl: Math.max(60, Math.floor((at.expires_at - Date.now()) / 1000)) });
  return at;
}

async function graphMe(accessToken) {
  const res = await fetch("https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName,mail", {
    headers: { authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) return null;
  const me = await res.json();
  return { id: me.id, email: me.mail || me.userPrincipalName || "" };
}

async function deviceKey(request) {
  const auth = request.headers.get("authorization") ?? "";
  const m = /^Bearer\s+(\S+)$/.exec(auth);
  return m ? `dev:${await sha256b64url(m[1])}` : null;
}

async function device(request, env) {
  const key = await deviceKey(request);
  return key ? env.STORE.get(key, "json") : null;
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });
}

function cors(env, res) {
  const h = new Headers(res.headers);
  h.set("access-control-allow-origin", env.APP_ORIGIN || "*");
  h.set("access-control-allow-methods", "GET, POST, OPTIONS");
  h.set("access-control-allow-headers", "authorization, content-type");
  h.set("access-control-max-age", "86400");
  h.set("vary", "origin");
  return new Response(res.body, { status: res.status, headers: h });
}

function randomToken(bytes) {
  const a = new Uint8Array(bytes);
  crypto.getRandomValues(a);
  return b64url(a);
}

async function sha256b64url(text) {
  const d = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return b64url(new Uint8Array(d));
}

function b64url(bytes) {
  let s = "";
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
