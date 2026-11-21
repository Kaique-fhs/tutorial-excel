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

  /* ---------- um modulo por tela ---------- */
  var links = $$("#menu a[href^='#']");
  var mapa = {};
  links.forEach(function (a) { mapa[a.getAttribute("href").slice(1)] = a; });
  var modulos = $$(".modulo, .capa");
  var listaModulos = $$(".modulo");
  raiz.classList.add("paginado");

  function nomeDe(m) {
    var a = mapa[m.id];
    return a ? a.lastChild.textContent.trim() : $("h2", m).textContent;
  }
  function numeroDe(m) {
    var a = mapa[m.id];
    return a ? $(".num", a).textContent.trim() : "";
  }
  function escapar(t) {
    return String(t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; });
  }

  // grade de cartoes na pagina inicial
  var vitrine = $("#vitrine .vitrine-conteudo");
  if (vitrine) {
    var html = "";
    $$("#menu .grupo").forEach(function (g) {
      var ol = g.nextElementSibling;
      html += '<h3 class="vitrine-grupo">' + escapar(g.textContent) + '</h3><div class="vitrine-grade">';
      $$("a", ol).forEach(function (a) {
        var id = a.getAttribute("href").slice(1);
        var m = document.getElementById(id);
        if (!m) return;
        var resumo = $(".objetivo li", m) || $(":scope > p", m);
        var chips = $$(".meta .chip", m).slice(0, 2).map(function (c) {
          return '<span class="' + c.className + '">' + escapar(c.textContent) + "</span>";
        }).join("");
        html += '<a class="cartao-modulo" href="#' + id + '" data-id="' + id + '">' +
          '<span class="cm-num">' + escapar(numeroDe(m)) + "</span>" +
          '<span class="cm-corpo"><span class="cm-nome">' + escapar(nomeDe(m)) + "</span>" +
          (resumo ? '<span class="cm-resumo">' + escapar(resumo.textContent) + "</span>" : "") +
          (chips ? '<span class="cm-chips">' + chips + "</span>" : "") +
          '</span><span class="cm-feito" aria-label="Concluído">✓</span></a>';
      });
      html += "</div>";
    });
    vitrine.innerHTML = html;
  }

  // em cada modulo: voltar, sumario e anterior/proximo
  listaModulos.forEach(function (m, i) {
    var cab = $("header", m);
    if (cab) {
      var voltar = document.createElement("a");
      voltar.className = "voltar-modulos";
      voltar.href = "#inicio";
      voltar.innerHTML = '<span aria-hidden="true">←</span> Todos os módulos';
      cab.insertAdjacentElement("beforebegin", voltar);
    }
    var h3s = $$("h3", m);
    if (h3s.length >= 2) {
      var sum = document.createElement("nav");
      sum.className = "sumario";
      sum.setAttribute("aria-label", "Neste módulo");
      sum.innerHTML = '<span class="sumario-rotulo">Neste módulo</span>';
      h3s.forEach(function (h) {
        var b = document.createElement("button");
        b.type = "button";
        b.textContent = h.textContent;
        b.addEventListener("click", function () { h.scrollIntoView({ behavior: "smooth", block: "start" }); });
        sum.appendChild(b);
      });
      (($(".objetivo", m)) || cab).insertAdjacentElement("afterend", sum);
    }
    var ant = listaModulos[i - 1];
    var prox = listaModulos[i + 1];
    var nav = document.createElement("nav");
    nav.className = "nav-modulo";
    nav.setAttribute("aria-label", "Navegação entre módulos");
    nav.innerHTML =
      (ant ? '<a class="nav-ant" href="#' + ant.id + '"><small>← Anterior</small><span>' + escapar(nomeDe(ant)) + "</span></a>"
           : '<a class="nav-ant" href="#inicio"><small>← Voltar</small><span>Página inicial</span></a>') +
      (prox ? '<a class="nav-prox" href="#' + prox.id + '"><small>Próximo →</small><span>' + escapar(nomeDe(prox)) + "</span></a>"
            : '<a class="nav-prox" href="#inicio"><small>Fim do tutorial</small><span>Voltar ao início</span></a>');
    m.appendChild(nav);
  });

  function rota() {
    var id = decodeURIComponent(location.hash.slice(1));
    var alvo = id ? document.getElementById(id) : null;
    var mod = alvo ? (alvo.classList.contains("modulo") ? alvo : alvo.closest(".modulo")) : null;
    listaModulos.forEach(function (m) { m.classList.toggle("atual", m === mod); });
    document.body.classList.toggle("vendo-modulo", !!mod);
    links.forEach(function (l) {
      var on = !!mod && l.getAttribute("href") === "#" + mod.id;
      l.classList.toggle("ativo", on);
      if (on) l.setAttribute("aria-current", "page"); else l.removeAttribute("aria-current");
    });
    if (mod && alvo !== mod) alvo.scrollIntoView();
    else window.scrollTo({ top: 0, behavior: "instant" });
    document.title = mod ? nomeDe(mod) + " · Excel para Recrutamento" : "Excel para Recrutamento";
    if (mod) {
      var h2 = $("h2", mod);
      h2.setAttribute("tabindex", "-1");
      h2.focus({ preventScroll: true });
    }
  }
  window.addEventListener("hashchange", rota);

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
      var cartao = $('.cartao-modulo[data-id="' + id + '"]');
      if (cartao) cartao.classList.toggle("feito", ok);
    });
    var continuar = $("#continuar");
    if (continuar) {
      var pendente = botoesConcluir.filter(function (b) { return !estado.feitos[b.getAttribute("data-modulo")]; })[0];
      if (!feitos) {
        continuar.textContent = "Começar pelo módulo 0";
        continuar.href = "#como-usar";
      } else if (pendente) {
        var mp = document.getElementById(pendente.getAttribute("data-modulo"));
        continuar.textContent = "Continuar: " + nomeDe(mp);
        continuar.href = "#" + mp.id;
      } else {
        continuar.textContent = "Tudo concluído! Revisar o checklist";
        continuar.href = "#checklist";
      }
    }
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

  rota();
})();
