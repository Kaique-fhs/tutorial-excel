/* Tutorial de Excel para Recrutamento - interacoes
   Sem dependencias. Tudo funciona mesmo sem localStorage (aba anonima). */
(function () {
  "use strict";

  var CHAVE = "tutorial-excel-rs:v1";

  function ler() {
    try { return JSON.parse(localStorage.getItem(CHAVE)) || {}; } catch (e) { return {}; }
  }
  function gravar(estado) {
    try { localStorage.setItem(CHAVE, JSON.stringify(estado)); } catch (e) { /* sem armazenamento */ }
  }
  var estado = ler();
  estado.feitos = estado.feitos || {};
  estado.checklist = estado.checklist || {};

  function $(sel, raiz) { return (raiz || document).querySelector(sel); }
  function $$(sel, raiz) { return Array.prototype.slice.call((raiz || document).querySelectorAll(sel)); }

  function normalizar(txt) {
    return (txt || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
  }

  /* ---------- aviso rapido ---------- */
  var toast = $(".toast");
  var timerToast;
  function avisar(msg) {
    if (!toast) return;
    toast.textContent = msg;
    toast.classList.add("visivel");
    clearTimeout(timerToast);
    timerToast = setTimeout(function () { toast.classList.remove("visivel"); }, 1800);
  }

  /* ---------- tema ---------- */
  var raiz = document.documentElement;
  if (estado.tema) raiz.setAttribute("data-theme", estado.tema);
  var botaoTema = $("#botao-tema");
  if (botaoTema) {
    botaoTema.addEventListener("click", function () {
      var atual = raiz.getAttribute("data-theme") ||
        (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
      var novo = atual === "dark" ? "light" : "dark";
      raiz.setAttribute("data-theme", novo);
      estado.tema = novo;
      gravar(estado);
      avisar(novo === "dark" ? "Tema escuro ativado" : "Tema claro ativado");
    });
  }

  /* ---------- menu lateral (celular) ---------- */
  var menu = $("#menu");
  var fundo = $(".fundo-menu");
  var botaoMenu = $("#botao-menu");
  function abrirMenu(abrir) {
    if (!menu) return;
    menu.classList.toggle("aberto", abrir);
    if (fundo) fundo.classList.toggle("aberto", abrir);
    if (botaoMenu) botaoMenu.setAttribute("aria-expanded", abrir ? "true" : "false");
  }
  if (botaoMenu) botaoMenu.addEventListener("click", function () { abrirMenu(!menu.classList.contains("aberto")); });
  if (fundo) fundo.addEventListener("click", function () { abrirMenu(false); });
  document.addEventListener("keydown", function (ev) { if (ev.key === "Escape") abrirMenu(false); });
  $$("#menu a").forEach(function (a) { a.addEventListener("click", function () { abrirMenu(false); }); });

  /* ---------- secao ativa no menu ---------- */
  var links = $$("#menu a[href^='#']");
  var mapa = {};
  links.forEach(function (a) { mapa[a.getAttribute("href").slice(1)] = a; });
  var modulos = $$(".modulo, .capa");
  if ("IntersectionObserver" in window) {
    var obs = new IntersectionObserver(function (entradas) {
      entradas.forEach(function (en) {
        if (en.isIntersecting && mapa[en.target.id]) {
          links.forEach(function (l) { l.classList.remove("ativo"); l.removeAttribute("aria-current"); });
          mapa[en.target.id].classList.add("ativo");
          mapa[en.target.id].setAttribute("aria-current", "true");
        }
      });
    }, { rootMargin: "-20% 0px -70% 0px" });
    modulos.forEach(function (m) { obs.observe(m); });
  }

  /* ---------- busca ---------- */
  var busca = $("#busca");
  var semResultado = $(".sem-resultado");
  var textos = {};
  modulos.forEach(function (m) { textos[m.id] = normalizar(m.textContent); });
  if (busca) {
    busca.addEventListener("input", function () {
      var q = normalizar(busca.value.trim());
      var visiveis = 0;
      links.forEach(function (a) {
        var id = a.getAttribute("href").slice(1);
        var bate = !q || normalizar(a.textContent).indexOf(q) > -1 || (textos[id] || "").indexOf(q) > -1;
        a.parentElement.hidden = !bate;
        if (bate) visiveis++;
      });
      $$("#menu .grupo").forEach(function (g) {
        var lista = g.nextElementSibling;
        var algum = lista && $$("li", lista).some(function (li) { return !li.hidden; });
        g.hidden = !algum;
      });
      if (semResultado) semResultado.hidden = visiveis > 0;
    });
  }

  /* ---------- progresso ---------- */
  var botoesConcluir = $$(".concluir");
  function atualizarProgresso() {
    var total = botoesConcluir.length;
    var feitos = 0;
    botoesConcluir.forEach(function (b) {
      var id = b.getAttribute("data-modulo");
      var ok = !!estado.feitos[id];
      if (ok) feitos++;
      b.setAttribute("aria-pressed", ok ? "true" : "false");
      b.querySelector(".texto").textContent = ok ? "Módulo concluído" : "Marcar como concluído";
      if (mapa[id]) mapa[id].classList.toggle("feito", ok);
    });
    var pct = total ? Math.round((feitos / total) * 100) : 0;
    var barra = $(".progresso-topo .barra span");
    var rotulo = $(".progresso-topo .valor");
    if (barra) barra.style.width = pct + "%";
    if (rotulo) rotulo.textContent = pct + "%";
    var prog = $(".progresso-topo");
    if (prog) prog.setAttribute("aria-label", "Progresso: " + feitos + " de " + total + " módulos concluídos");
  }
  botoesConcluir.forEach(function (b) {
    b.addEventListener("click", function () {
      var id = b.getAttribute("data-modulo");
      estado.feitos[id] = !estado.feitos[id];
      gravar(estado);
      atualizarProgresso();
      if (estado.feitos[id]) avisar("Boa! Módulo concluído");
    });
  });
  var zerar = $("#zerar-progresso");
  if (zerar) zerar.addEventListener("click", function () {
    estado.feitos = {};
    estado.checklist = {};
    gravar(estado);
    $$(".checklist input").forEach(function (c) { c.checked = false; });
    atualizarProgresso();
    avisar("Progresso zerado");
  });
  atualizarProgresso();

  /* ---------- copiar formula ---------- */
  function copiarTexto(txt) {
    if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(txt);
    return new Promise(function (ok, falha) {
      var area = document.createElement("textarea");
      area.value = txt;
      area.setAttribute("readonly", "");
      area.style.position = "fixed";
      area.style.opacity = "0";
      document.body.appendChild(area);
      area.select();
      try { document.execCommand("copy") ? ok() : falha(); } catch (e) { falha(e); }
      document.body.removeChild(area);
    });
  }
  $$(".formula").forEach(function (bloco) {
    var cab = $(".cab", bloco);
    var pre = $("pre", bloco);
    if (!cab || !pre) return;
    var b = document.createElement("button");
    b.type = "button";
    b.className = "copiar";
    b.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg><span>Copiar</span>';
    b.setAttribute("aria-label", "Copiar fórmula");
    b.addEventListener("click", function () {
      copiarTexto(pre.textContent.trim()).then(function () {
        b.querySelector("span").textContent = "Copiado!";
        avisar("Fórmula copiada. Cole no Excel.");
        setTimeout(function () { b.querySelector("span").textContent = "Copiar"; }, 1500);
      }, function () { avisar("Não deu para copiar. Selecione o texto e copie."); });
    });
    cab.appendChild(b);
  });

  /* ---------- quiz ---------- */
  $$(".quiz").forEach(function (q) {
    var correta = parseInt(q.getAttribute("data-correta"), 10);
    var retorno = $(".retorno", q);
    var opcoes = $$("button.opcao", q);
    opcoes.forEach(function (b, i) {
      b.addEventListener("click", function () {
        opcoes.forEach(function (o) { o.classList.remove("correta", "errada"); });
        if (i === correta) {
          b.classList.add("correta");
          retorno.innerHTML = "<strong>Correto.</strong> " + (q.getAttribute("data-explica") || "");
        } else {
          b.classList.add("errada");
          retorno.innerHTML = "<strong>Ainda não.</strong> Releia o trecho acima e tente de novo.";
        }
      });
    });
  });

  /* ---------- calculadora de funil ---------- */
  var calc = $("#calc-funil");
  if (calc) {
    var campos = $$("input", calc);
    var saida = $("#saida-funil");
    var extras = $("#saida-extras");
    function fmt(n) { return isFinite(n) ? (n * 100).toFixed(1).replace(".", ",") + "%" : "–"; }
    function calcular() {
      var vals = campos.map(function (c) { return Math.max(0, parseFloat(c.value) || 0); });
      var nomes = campos.map(function (c) { return c.getAttribute("data-nome"); });
      var topo = vals[0] || 1;
      var html = "";
      for (var i = 0; i < 5; i++) {
        var larg = Math.min(100, (vals[i] / topo) * 100);
        var taxa = i === 0 ? "100%" : fmt(vals[i] / vals[i - 1]);
        html += '<div class="linha"><span>' + nomes[i] + '</span><div class="trilho"><div class="preenche" style="width:' +
          larg + '%">' + vals[i] + '</div></div><span class="taxa">' + taxa + "</span></div>";
      }
      saida.innerHTML = html;
      var custo = vals[6];
      extras.innerHTML =
        '<div class="cartao"><h4>Aceite de proposta</h4><p class="grande">' + fmt(vals[4] / vals[3]) + "</p><p>contratados ÷ propostas</p></div>" +
        '<div class="cartao"><h4>Conversão total</h4><p class="grande">' + fmt(vals[4] / vals[0]) + "</p><p>contratados ÷ inscritos</p></div>" +
        '<div class="cartao"><h4>Inscritos por vaga</h4><p class="grande">' + (vals[5] ? (vals[0] / vals[5]).toFixed(1).replace(".", ",") : "–") + "</p><p>inscritos ÷ vagas</p></div>" +
        '<div class="cartao"><h4>Custo por contratação</h4><p class="grande">' + (vals[4] ? (custo / vals[4]).toLocaleString("pt-BR", { style: "currency", currency: "BRL" }) : "–") + "</p><p>custo total ÷ contratados</p></div>";
    }
    campos.forEach(function (c) { c.addEventListener("input", calcular); });
    calcular();
  }

  /* ---------- checklist ---------- */
  $$(".checklist input").forEach(function (c, i) {
    var id = c.id || "item-" + i;
    c.checked = !!estado.checklist[id];
    c.addEventListener("change", function () {
      estado.checklist[id] = c.checked;
      gravar(estado);
    });
  });

  /* ---------- voltar ao topo ---------- */
  var voltar = $(".voltar-topo");
  if (voltar) {
    window.addEventListener("scroll", function () {
      voltar.classList.toggle("visivel", window.scrollY > 900);
    }, { passive: true });
    voltar.addEventListener("click", function () { window.scrollTo({ top: 0 }); });
  }

  /* ---------- ano no rodape ---------- */
  var ano = $("#ano");
  if (ano) ano.textContent = new Date().getFullYear();
})();
