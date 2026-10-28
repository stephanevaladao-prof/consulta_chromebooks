/* Consulta de Chromebook: lê a planilha do inventário na mesma pasta e mostra os dados de um N/S. */
(function () {
  'use strict';

  var ARQUIVO = 'Inventario_Carrinhos_conferido.xlsx';
  try {
    var pedido = new URLSearchParams(location.search).get('arquivo');
    if (pedido && /^[\w .()\-]+\.xlsx$/i.test(pedido)) ARQUIVO = pedido;
  } catch (e) { /* usa o nome padrão */ }
  var LS_KEY = 'consulta-recentes';
  var MAX = 30;
  var $ = function (s) { return document.querySelector(s); };

  var D = null;

  /* ---------- utilidades ---------- */
  function txt(v) { return v == null ? '' : String(v).trim(); }
  function norm(s) { return txt(s).toUpperCase().replace(/[^A-Z0-9]/g, ''); }
  function keyOf(sn) { var n = norm(sn); return n.indexOf('09S') === 0 ? n.slice(0, 14) : n; }
  function semAcento(t) { return String(t).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase(); }
  function h(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }
  function cartId(nome) {
    var t = semAcento(nome);
    if (t.indexOf('reserva') >= 0) return 'res';
    var m = t.match(/carrinho\s*0*(\d+)/);
    return m ? 'c' + m[1] : t;
  }
  var CORES = ['azul', 'amarelo', 'vermelho', 'laranja', 'roxo', 'verde', 'rosa', 'cinza', 'marrom', 'preto', 'ocre', 'ciano'];
  function cartInfo(nome) {
    var id = cartId(nome), t = semAcento(nome);
    if (id === 'res') return { id: id, titulo: 'Reserva', cor: '' };
    var m = t.match(/carrinho\s*0*(\d+)/);
    var cor = CORES.filter(function (c) { return t.indexOf(c) >= 0; })[0] || '';
    var num = m ? ('0' + m[1]).slice(-2) : txt(nome);
    return { id: id, titulo: 'Carrinho ' + num, cor: cor + (m && m[1] === '13' ? (cor ? ', ' : '') + 'quebrados' : '') };
  }

  function modeloSn(sn, padrao) {
    var n = norm(sn);
    if (n.indexOf('09S29Q') === 0) return 'Samsung Chromebook 4 prata';
    if (n.indexOf('NXG54AL') === 0) return 'Acer Chromebook R11 CB5-132T branco';
    if (n.indexOf('NXH0HAL0018') === 0) return 'Acer Chromebook 11 C731 preto';
    return padrao || '';
  }
  var RE_FORA = /n[ãa]o liga|n[ãa]o est[áa] ligando|n[ãa]o carrega|tela quebrada|imposs[íi]vel de usar|totalmente quebrado|com defeito|n[ãa]o est[áa] pr[óo]prio|problema na tela|tela piscando sozinha|tela est[áa] com problema|tremendo/i;
  var RE_RESSALVA = /dobradi[çc]a|abre e fecha|soltando|v[áa]rias teclas|n[ãa]o est[ãa]o funcionando|power|piscando/i;
  function uso(t) { if (RE_FORA.test(t)) return 'fora'; if (RE_RESSALVA.test(t)) return 'ressalva'; return 'ok'; }

  /* ---------- leitura da planilha ---------- */
  function rowsOf(wb, name) {
    var achada = wb.SheetNames.filter(function (n) { return semAcento(n) === semAcento(name); })[0];
    if (!achada) throw new Error('A planilha não tem a aba "' + name + '". Abas encontradas: ' + (wb.SheetNames.join(', ') || 'nenhuma') + '. Escolha o arquivo Inventario_Carrinhos_conferido.xlsx.');
    return XLSX.utils.sheet_to_json(wb.Sheets[achada], { header: 1, defval: null, raw: true });
  }
  function rowsOpt(wb, name) {
    try { return rowsOf(wb, name); } catch (e) { return []; }
  }

  function parse(wb) {
    var cf = rowsOf(wb, 'Conferência'), sy = rowsOf(wb, 'Sistema'), dv = rowsOpt(wb, 'Divergências'), co = rowsOpt(wb, 'Comodato');
    var words = [], sys = [], com = [], cases = [], i, r, ok;

    ok = false;
    for (i = 0; i < cf.length; i++) {
      r = cf[i];
      if (!ok) { if (txt(r[0]) === 'Carrinho (Word)') ok = true; continue; }
      if (!txt(r[0]) || !txt(r[2])) continue;
      words.push({ cart: txt(r[0]), n: txt(r[1]), sn: txt(r[2]), key: keyOf(txt(r[3]) || txt(r[2])), pat: txt(r[5]), modelo: txt(r[6]),
        locSis: txt(r[7]), modeloSis: txt(r[8]), sit: txt(r[10]), obs: txt(r[11]), estado: txt(r[12]) });
    }
    if (!words.length) throw new Error('Não encontrei a lista de computadores na aba "Conferência". Confira se é a planilha certa.');

    ok = false;
    for (i = 0; i < sy.length; i++) {
      r = sy[i];
      if (!ok) { if (txt(r[3]).indexOf('N. S') === 0) ok = true; continue; }
      if (!txt(r[3])) continue;
      sys.push({ sn: txt(r[3]), key: keyOf(txt(r[6]) || txt(r[3])), modelo: txt(r[2]), sit: txt(r[4]), local: txt(r[5]), conf: txt(r[7]), obs: txt(r[8]) });
    }

    ok = false;
    for (i = 0; i < co.length; i++) {
      r = co[i];
      if (!ok) { if (txt(r[0]) === 'RM/RSE') ok = true; continue; }
      if (!txt(r[4])) continue;
      com.push({ cargo: txt(r[2]), modelo: txt(r[3]), sn: txt(r[4]), key: keyOf(txt(r[4])), obs: txt(r[7]) });
    }

    for (i = 0; i < dv.length; i++) {
      r = dv[i];
      var c0 = txt(r[0]);
      if (c0.indexOf('Para aten') === 0) break;
      if (!c0 || isNaN(Number(c0))) continue;
      cases.push({ n: Number(c0), cart: txt(r[1]), item: txt(r[2]), tipo: txt(r[3]), snWord: txt(r[4]), snSis: txt(r[6]),
        porque: txt(r[8]), fazer: txt(r[9]), ver: txt(r[10]).toLowerCase() === 'sim', res: txt(r[11]) });
    }
    return { words: words, sys: sys, com: com, cases: cases };
  }

  /* ---------- busca ---------- */
  function casa(sn, key, q, qKey) {
    var n = norm(sn);
    if (n.indexOf(q) >= 0) return true;
    return !!qKey && qKey.length >= 14 && key === qKey;
  }

  function buscar(raw) {
    var q = norm(raw);
    if (q.length < 4) return null;
    var qKey = q.indexOf('09S') === 0 ? q.slice(0, 14) : '';
    var soNum = /^\d+$/.test(q);

    var w = D.words.filter(function (x) { return norm(x.sn).indexOf(q) >= 0 || (soNum && norm(x.pat) === q); });
    var s = D.sys.filter(function (x) { return casa(x.sn, x.key, q, qKey); });
    var c = D.com.filter(function (x) { return casa(x.sn, x.key, q, qKey); });

    var jaW = {};
    w.forEach(function (x) { jaW[x.cart + '|' + x.n] = 1; });
    var chavesDiretas = {};
    w.forEach(function (x) { chavesDiretas[x.key] = 1; });
    s.forEach(function (x) {
      if (chavesDiretas[x.key]) return;
      D.words.forEach(function (y) { if (y.key === x.key && !jaW[y.cart + '|' + y.n]) { w.push(y); jaW[y.cart + '|' + y.n] = 1; } });
    });
    var chavesW = {};
    D.words.forEach(function (x) { chavesW[x.key] = 1; });
    var soSis = s.filter(function (x) { return !chavesW[x.key]; });
    return { words: w, soSis: soSis, com: c };
  }

  function sysDe(w) {
    var l = D.sys.filter(function (x) { return x.key === w.key; });
    if (!l.length) return null;
    var exato = l.filter(function (x) { return norm(x.sn) === norm(w.sn); })[0];
    return exato || l[0];
  }
  function casosDe(w, sis, cartNome, item) {
    var id = cartNome ? cartId(cartNome) : '';
    var key = w ? w.key : sis ? sis.key : '';
    return D.cases.filter(function (c) {
      if (w && item && cartId(c.cart) === id && c.item === item) return true;
      return !!key && c.snSis && keyOf(c.snSis) === key;
    });
  }

  /* ---------- desenho ---------- */
  function localBonito(l) { if (!txt(l)) return ''; var i = cartInfo(l); return i.titulo + (i.cor ? ' (' + i.cor + ')' : ''); }
  function linha(k, v, opt) {
    opt = opt || {};
    var r = h('div', 'row');
    r.appendChild(h('span', 'k', k));
    var vazio = !v;
    r.appendChild(h('span', 'v' + (opt.mono ? ' mono' : '') + (opt.diff ? ' diff' : '') + (vazio ? ' none' : ''), vazio ? (opt.vazio || 'não informado') : v));
    return r;
  }
  function badge(cls, t) { return h('span', 'badge ' + cls, t); }
  function cabecalho(k, v, m, b) {
    var hd = h('div', 'card-h'), wh = h('div', 'where');
    wh.appendChild(h('span', 'k', k)); wh.appendChild(h('span', 'v', v));
    if (m) wh.appendChild(h('span', 'm', m));
    hd.appendChild(wh); hd.appendChild(b);
    return hd;
  }
  function veredito(cls, titulo, texto) {
    var v = h('div', 'verd ' + cls);
    v.appendChild(h('div', 't', titulo));
    if (texto) v.appendChild(h('p', null, texto));
    return v;
  }
  function blocoCaso(c) {
    var b = h('div', 'case');
    b.appendChild(h('div', 't', (c.ver ? 'Caso resolvido: ' : 'Caso em aberto: ') + c.tipo));
    if (c.ver && c.res) { b.appendChild(h('div', 'lbl', 'Resultado')); b.appendChild(h('p', null, c.res)); }
    if (c.porque) { b.appendChild(h('div', 'lbl', 'O que provavelmente aconteceu')); b.appendChild(h('p', null, c.porque)); }
    if (!c.ver && c.fazer) { b.appendChild(h('div', 'lbl', 'O que conferir')); b.appendChild(h('p', null, c.fazer)); }
    return b;
  }

  function cardWord(w) {
    var ci = cartInfo(w.cart), sis = sysDe(w), u = uso(w.estado);
    var card = h('article', 'card');
    var bd = u === 'fora' ? badge('b-bad', 'Fora de uso') : u === 'ressalva' ? badge('b-warn', 'Usável, com defeito') : badge('b-ok', 'Usável');
    card.appendChild(cabecalho('Onde está', ci.titulo + ' · nº ' + w.n, [ci.cor ? 'carrinho ' + ci.cor : '', modeloSn(w.sn, w.modelo)].filter(Boolean).join(' · '), bd));

    var cols = h('div', 'cols'), a = h('div', 'col'), b = h('div', 'col');
    a.appendChild(h('h3', null, 'O que eu anotei (Word)'));
    var difSn = sis && norm(sis.sn) !== norm(w.sn);
    a.appendChild(linha('N/S ditado', w.sn, { mono: true, diff: difSn }));
    a.appendChild(linha('Patrimônio', w.pat, { vazio: 'não localizado' }));
    a.appendChild(linha('Modelo', modeloSn(w.sn, w.modelo)));
    a.appendChild(linha('Estado informado', w.estado, { vazio: 'nenhum defeito anotado' }));

    b.appendChild(h('h3', null, 'O que o sistema diz'));
    if (sis) {
      var difLoc = cartId(sis.local) !== cartId(w.cart);
      b.appendChild(linha('N. Série', sis.sn, { mono: true, diff: difSn }));
      b.appendChild(linha('Local', localBonito(sis.local), { diff: difLoc }));
      b.appendChild(linha('Situação', sis.sit));
      b.appendChild(linha('Modelo no cadastro', sis.modelo));
    } else {
      b.appendChild(linha('N. Série', '', { vazio: 'não existe no sistema' }));
      b.appendChild(linha('Local', '', { vazio: '—' }));
    }
    cols.appendChild(a); cols.appendChild(b); card.appendChild(cols);

    var casos = casosDe(w, sis, w.cart, w.n);
    var abertos = casos.filter(function (c) { return !c.ver; });
    if (w.sit === 'Confere' && !abertos.length) {
      card.appendChild(veredito('ok', 'Confere com o sistema', 'O N/S e o local do Word batem com o cadastro.'));
    } else {
      card.appendChild(veredito(abertos.length ? 'warn' : 'ok', abertos.length ? 'Precisa de conferência' : 'Resolvido', w.obs || w.sit));
    }
    casos.forEach(function (c) { card.appendChild(blocoCaso(c)); });
    return card;
  }

  function cardSoSis(s) {
    var ci = cartInfo(s.local), card = h('article', 'card');
    card.appendChild(cabecalho('Cadastrado no sistema em', ci.titulo, modeloSn(s.sn, s.modelo), badge('b-info', 'Não apareceu no Word')));
    var cols = h('div', 'cols'), a = h('div', 'col'), b = h('div', 'col');
    a.appendChild(h('h3', null, 'O que eu anotei (Word)'));
    a.appendChild(linha('N/S ditado', '', { vazio: 'não apareceu em nenhum Word' }));
    b.appendChild(h('h3', null, 'O que o sistema diz'));
    b.appendChild(linha('N. Série', s.sn, { mono: true }));
    b.appendChild(linha('Local', localBonito(s.local)));
    b.appendChild(linha('Situação', s.sit));
    b.appendChild(linha('Modelo no cadastro', s.modelo));
    cols.appendChild(a); cols.appendChild(b); card.appendChild(cols);
    var cs = casosDe(null, s, '', ''), abertosS = cs.filter(function (c) { return !c.ver; }), resolvidos = cs.filter(function (c) { return c.ver; });
    if (resolvidos.length && !abertosS.length) {
      card.appendChild(veredito('ok', 'Caso resolvido', resolvidos[0].res || s.obs));
      card.querySelector('.badge').className = 'badge b-ok'; card.querySelector('.badge').textContent = 'Resolvido';
    } else {
      card.appendChild(veredito('warn', 'Equipamento a procurar', s.obs || 'Está no sistema, mas nenhum Word tem este N/S. Pode estar com alguém, em outro carrinho ou na manutenção.'));
    }
    cs.forEach(function (c) { card.appendChild(blocoCaso(c)); });
    return card;
  }

  function cardCom(c) {
    var card = h('article', 'card');
    card.appendChild(cabecalho('Situação', 'Em comodato', modeloSn(c.sn, c.modelo), badge('b-info', 'Comodato')));
    var cols = h('div', 'cols'), a = h('div', 'col'), b = h('div', 'col');
    a.appendChild(h('h3', null, 'Lista de comodato'));
    a.appendChild(linha('N/S', c.sn, { mono: true }));
    a.appendChild(linha('Cargo de quem usa', c.cargo));
    b.appendChild(h('h3', null, 'Sistema da escola'));
    b.appendChild(linha('Situação', '', { vazio: 'fora do inventário da escola' }));
    cols.appendChild(a); cols.appendChild(b); card.appendChild(cols);
    card.appendChild(veredito('info', 'Aparelho em comodato', c.obs || 'Este aparelho está com um professor ou funcionário. Por isso não está nos carrinhos nem no cadastro do local da escola.'));
    return card;
  }

  function mostrar(raw) {
    var out = $('#results'), st = $('#status');
    out.textContent = '';
    var res = buscar(raw);
    if (res === null) { st.textContent = txt(raw) ? 'Digite pelo menos 4 caracteres.' : D.words.length + ' computadores no Word e ' + D.sys.length + ' no sistema.'; return; }
    var total = res.words.length + res.soSis.length + res.com.length;
    if (!total) {
      st.textContent = 'Nenhum resultado.';
      var e = h('div', 'empty');
      e.appendChild(h('b', null, 'Não achei "' + txt(raw).toUpperCase() + '"'));
      e.appendChild(h('span', 'muted', 'Confira se não ficou nenhuma letra trocada (0 e O, 1 e I, 8 e B costumam se confundir). Dá para digitar só um pedaço do número. Se o aparelho for novo, ele pode não estar nem no Word nem no sistema.'));
      out.appendChild(e);
      return;
    }
    st.textContent = total + (total === 1 ? ' resultado' : ' resultados');
    var n = 0;
    function add(card) { if (n < MAX) { out.appendChild(card); } n++; }
    res.words.forEach(function (x) { add(cardWord(x)); });
    res.soSis.forEach(function (x) { add(cardSoSis(x)); });
    res.com.forEach(function (x) { add(cardCom(x)); });
    if (n > MAX) out.appendChild(h('p', 'more muted', 'Mostrando os primeiros ' + MAX + ' de ' + n + '. Digite mais caracteres para afunilar.'));
    guardar(raw);
  }

  /* ---------- buscas recentes ---------- */
  function lerRecentes() { try { return JSON.parse(localStorage.getItem(LS_KEY) || '[]'); } catch (e) { return []; } }
  function guardar(raw) {
    var v = txt(raw).toUpperCase();
    if (v.length < 4) return;
    var l = lerRecentes().filter(function (x) { return x !== v; });
    l.unshift(v); l = l.slice(0, 5);
    try { localStorage.setItem(LS_KEY, JSON.stringify(l)); } catch (e) { /* sem armazenamento */ }
    desenharRecentes();
  }
  function desenharRecentes() {
    var box = $('#recent'), l = lerRecentes();
    box.textContent = '';
    if (!l.length) { box.hidden = true; return; }
    box.hidden = false;
    box.appendChild(h('span', null, 'Últimas:'));
    l.forEach(function (v) {
      var b = h('button', null, v); b.type = 'button';
      b.addEventListener('click', function () { $('#q').value = v; mostrar(v); });
      box.appendChild(b);
    });
  }

  /* ---------- carregamento ---------- */
  function erro(html) { var m = $('#msg'); m.textContent = ''; m.appendChild(html); m.hidden = false; }
  function msgErro(titulo, detalhe) {
    var d = document.createElement('div');
    var b = h('b', null, titulo); d.appendChild(b); d.appendChild(document.createElement('br')); d.appendChild(document.createTextNode(detalhe));
    erro(d);
  }
  function pronto(wb, origem) {
    try { D = parse(wb); } catch (e) { msgErro('Não consegui ler esta planilha.', e.message); $('#status').textContent = ''; return; }
    $('#msg').hidden = true;
    $('#q').disabled = false; $('#btn-go').disabled = false;
    $('#src-info').textContent = 'Planilha: ' + origem + ' · ' + D.words.length + ' computadores';
    desenharRecentes();
    mostrar($('#q').value);
    if (!$('#q').value) $('#q').focus({ preventScroll: true });
  }
  function carregar() {
    if (typeof XLSX === 'undefined') { msgErro('Não foi possível carregar a biblioteca de leitura de planilhas.', 'Confira se o arquivo xlsx.full.min.js está na mesma pasta desta página.'); return; }
    $('#status').textContent = 'Carregando a planilha…';
    fetch(ARQUIVO + '?t=' + Date.now()).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.arrayBuffer();
    }).then(function (buf) {
      pronto(XLSX.read(buf, { type: 'array' }), ARQUIVO);
    }).catch(function () {
      $('#status').textContent = '';
      msgErro('Não achei a planilha ' + ARQUIVO + ' ao lado desta página.', 'Coloque o arquivo na mesma pasta (no GitHub, na mesma lista de arquivos) ou use "Escolher planilha" lá embaixo.');
    });
  }

  $('#file-in').addEventListener('change', function (ev) {
    var f = ev.target.files && ev.target.files[0];
    if (!f) return;
    var rd = new FileReader();
    rd.onload = function () { pronto(XLSX.read(rd.result, { type: 'array' }), f.name); };
    rd.readAsArrayBuffer(f);
  });
  $('#btn-reload').addEventListener('click', carregar);
  $('#form').addEventListener('submit', function (ev) { ev.preventDefault(); if (D) { mostrar($('#q').value); $('#q').blur(); } });
  $('#q').addEventListener('input', function () { if (D) mostrar(this.value); });

  /* ---------- leitor de código de barras (quando o navegador permite) ---------- */
  var stream = null, rodando = false;
  function fecharScan() {
    rodando = false;
    if (stream) { stream.getTracks().forEach(function (t) { t.stop(); }); stream = null; }
    $('#scan').hidden = true;
  }
  if ('BarcodeDetector' in window && navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
    $('#btn-scan').hidden = false;
    $('#btn-scan').addEventListener('click', function () {
      var det;
      try { det = new BarcodeDetector(); } catch (e) { det = null; }
      if (!det) return;
      navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' }, audio: false }).then(function (s) {
        stream = s; rodando = true;
        var v = $('#video'); v.srcObject = s; v.play();
        $('#scan').hidden = false;
        $('#scan-msg').textContent = 'Aponte a câmera para o código de barras da etiqueta.';
        (function volta() {
          if (!rodando) return;
          det.detect(v).then(function (codes) {
            if (codes && codes.length) {
              var val = codes[0].rawValue;
              fecharScan();
              $('#q').value = val; mostrar(val);
            } else { setTimeout(volta, 250); }
          }).catch(function () { setTimeout(volta, 400); });
        })();
      }).catch(function () { alert('Não consegui abrir a câmera. Confira se o navegador tem permissão para usá-la.'); });
    });
    $('#scan-close').addEventListener('click', fecharScan);
  }

  carregar();
})();
