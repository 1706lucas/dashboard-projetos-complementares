// Cofre: login e criptografia dos dados no navegador (AES-GCM 256 + PBKDF2-SHA256).
// Os arquivos de dados ficam publicados como <nome>.enc. Só quem tem login consegue abri-los.
(function () {
  var ITER = 600000;
  var SKEY = "gadens.sessao.v1";
  var USUARIOS = "data/usuarios.json";
  var enc = new TextEncoder(), dec = new TextDecoder();

  function b64(buf) {
    var b = new Uint8Array(buf), s = "";
    for (var i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
    return btoa(s);
  }
  function unb64(s) {
    var bin = atob(s), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function rand(n) { return crypto.getRandomValues(new Uint8Array(n)); }
  function norm(u) { return String(u || "").trim().toLowerCase(); }

  async function chaveSenha(senha, salt, iter) {
    var base = await crypto.subtle.importKey("raw", enc.encode(senha), "PBKDF2", false, ["deriveKey"]);
    return crypto.subtle.deriveKey({ name: "PBKDF2", hash: "SHA-256", salt: salt, iterations: iter },
      base, { name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
  }
  function importaDK(raw) {
    return crypto.subtle.importKey("raw", raw, { name: "AES-GCM" }, true, ["encrypt", "decrypt"]);
  }
  async function cifraBytes(key, bytes) {
    var iv = rand(12);
    var ct = await crypto.subtle.encrypt({ name: "AES-GCM", iv: iv }, key, bytes);
    return { v: 1, iv: b64(iv), ct: b64(ct) };
  }
  async function decifraBytes(key, pkt) {
    return new Uint8Array(await crypto.subtle.decrypt({ name: "AES-GCM", iv: unb64(pkt.iv) }, key, unb64(pkt.ct)));
  }

  var cacheUsuarios = null;
  async function usuarios() {
    if (cacheUsuarios) return cacheUsuarios;
    try {
      var r = await fetch(USUARIOS, { cache: "no-store" });
      if (!r.ok) return null;
      var j = await r.json();
      cacheUsuarios = j && Array.isArray(j.usuarios) ? j : null;
    } catch (e) { cacheUsuarios = null; }
    return cacheUsuarios;
  }

  var dkPromise = null;
  function sessao() {
    try { var s = JSON.parse(sessionStorage.getItem(SKEY) || "null"); return s && s.dk ? s : null; }
    catch (e) { return null; }
  }
  function dk() {
    var s = sessao(); if (!s) return Promise.resolve(null);
    if (!dkPromise) dkPromise = importaDK(unb64(s.dk));
    return dkPromise;
  }

  var Cofre = {
    // "aberto": sem usuarios.json (servidor interno) · "protegido": exige login
    modo: async function () { return (await usuarios()) ? "protegido" : "aberto"; },
    sessao: sessao,
    entrar: async function (login, senha) {
      var u = await usuarios();
      if (!u) throw new Error("Cadastro de usuários não encontrado.");
      var e = u.usuarios.find(function (x) { return x.login === norm(login); });
      if (!e) throw new Error("Login ou senha incorretos.");
      var k = await chaveSenha(senha, unb64(e.salt), e.iter || ITER);
      var raw;
      try { raw = await decifraBytes(k, e.chave); } catch (err) { throw new Error("Login ou senha incorretos."); }
      sessionStorage.setItem(SKEY, JSON.stringify({ login: e.login, nome: e.nome || e.login, perfil: e.perfil || "leitura", dk: b64(raw) }));
      dkPromise = null;
      return sessao();
    },
    sair: function () { sessionStorage.removeItem(SKEY); dkPromise = null; },
    // lê <caminho>.enc se existir; senão o arquivo aberto (modo servidor interno)
    carregar: async function (caminho) {
      var key = await dk();
      if (key) {
        var r = await fetch(caminho + ".enc", { cache: "no-store" });
        if (r.ok) return JSON.parse(dec.decode(await decifraBytes(key, await r.json())));
        if ((await Cofre.modo()) === "protegido") throw new Error("Arquivo " + caminho + ".enc não encontrado.");
      }
      var r2 = await fetch(caminho, { cache: "no-store" });
      if (!r2.ok) throw new Error("Falha ao carregar " + caminho);
      return r2.json();
    },
    cifrarJSON: async function (obj) {
      var key = await dk(); if (!key) throw new Error("Sessão expirada. Entre novamente.");
      return cifraBytes(key, enc.encode(JSON.stringify(obj)));
    },
    listaUsuarios: async function () { var u = await usuarios(); return u ? u.usuarios.slice() : []; },
    // cria a entrada de um usuário embrulhando a chave dos dados com a senha dele
    novoUsuario: async function (login, nome, perfil, senha) {
      var key = await dk(); if (!key) throw new Error("Sessão expirada. Entre novamente.");
      if (String(senha || "").length < 8) throw new Error("A senha precisa ter pelo menos 8 caracteres.");
      var raw = new Uint8Array(await crypto.subtle.exportKey("raw", key));
      var salt = rand(16);
      var k = await chaveSenha(senha, salt, ITER);
      return { login: norm(login), nome: String(nome || login).trim(), perfil: perfil === "admin" ? "admin" : "leitura",
        salt: b64(salt), iter: ITER, chave: await cifraBytes(k, raw), criado: new Date().toISOString().slice(0, 10) };
    },
    // configuração inicial: gera uma chave nova (usado só pela ferramenta de preparo)
    _inicial: async function (login, nome, senha) {
      var raw = rand(32); var key = await importaDK(raw);
      var salt = rand(16); var k = await chaveSenha(senha, salt, ITER);
      var entrada = { login: norm(login), nome: nome, perfil: "admin", salt: b64(salt), iter: ITER,
        chave: await cifraBytes(k, raw), criado: new Date().toISOString().slice(0, 10) };
      return { key: key, raw: b64(raw), entrada: entrada,
        cifrar: function (obj) { return cifraBytes(key, enc.encode(JSON.stringify(obj))); } };
    }
  };
  window.Cofre = Cofre;
})();
