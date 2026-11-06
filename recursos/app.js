/*
 * Assistente de Réplica Fiscal — Impugnação de IPTU
 * Interface: lista de processos, assistente em 4 etapas, prévia e emissão da réplica.
 */
(function () {
  'use strict';

  var T = window.Textos, P = window.Pasta, D = window.Docx, M = window.MODELO_REPLICA;

  var FORMATO = 'impugnacao-iptu/processo';
  var VERSAO = 1;
  var ARQ_CONFIG = '_configuracao-assistente.json';
  var CHAVE_LOCAL = 'assistente-impugnacao-iptu.config';
  var FISCAL_PADRAO = { nome: '', cargo: 'Auditora Fiscal da Receita Municipal', matricula: '91979' };

  var SUGESTOES = [
    { rotulo: 'Retificação cadastral', texto: 'Recomenda-se, por conseguinte, a retificação do cadastro imobiliário do CCI nº {cci}, com a consequente revisão do lançamento do IPTU referente ao exercício de {exercicio}.' },
    { rotulo: 'Ausência de provas', texto: 'Ressalta-se que não foram apresentados, pelo(a) impugnante, elementos capazes de afastar os dados constantes do cadastro imobiliário municipal.' },
    { rotulo: 'Vistoria in loco', texto: 'Registra-se que a área construída foi aferida pelo(a) Técnico(a) Fiscal mediante vistoria in loco no imóvel.' }
  ];

  var estado = {
    processos: [],          // [{arquivo, dados, modificado}]
    invalidos: [],          // [{arquivo, erro}]
    atual: null,
    arquivoAtual: null,
    sujo: false,
    salvoEm: null,
    etapa: 0,
    config: { fiscal: clonar(FISCAL_PADRAO) },
    pasta: 'nenhuma',       // nenhuma | conectada | reconectar | download
    nomePastaDownload: '',
    imagens: { brasao: '', marca: '' },
    ultimaReplica: null
  };

  // ---------------------------------------------------------------- utilidades
  function $(s, r) { return (r || document).querySelector(s); }
  function $$(s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); }
  function clonar(o) { return JSON.parse(JSON.stringify(o)); }

  function esc(t) {
    return String(t === undefined || t === null ? '' : t)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function hoje() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  function agoraISO() { return new Date().toISOString(); }

  function dataHora(iso) {
    if (!iso) return '';
    var d = new Date(iso);
    if (isNaN(d)) return '';
    return d.toLocaleDateString('pt-BR') + ' ' + d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  }

  function gerarId() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    return 'p-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10);
  }

  function obter(obj, caminho) {
    return caminho.split('.').reduce(function (o, k) { return o == null ? undefined : o[k]; }, obj);
  }

  function definir(obj, caminho, valor) {
    var ks = caminho.split('.');
    var o = obj;
    for (var i = 0; i < ks.length - 1; i++) {
      if (typeof o[ks[i]] !== 'object' || o[ks[i]] === null) o[ks[i]] = {};
      o = o[ks[i]];
    }
    o[ks[ks.length - 1]] = valor;
  }

  // Mescla "dados" sobre "base" (objetos aninhados), preservando campos extras de "dados".
  function mesclar(base, dados) {
    if (!dados || typeof dados !== 'object' || Array.isArray(dados)) return base;
    Object.keys(dados).forEach(function (k) {
      var v = dados[k];
      if (v && typeof v === 'object' && !Array.isArray(v) && base[k] && typeof base[k] === 'object' && !Array.isArray(base[k])) {
        base[k] = mesclar(base[k], v);
      } else if (v !== undefined) {
        base[k] = v;
      }
    });
    return base;
  }

  function semAcento(t) {
    return String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '');
  }

  function seguro(t) {
    return semAcento(t).replace(/[^A-Za-z0-9-]+/g, '_').replace(/^_+|_+$/g, '');
  }

  function nomeArquivo(d) {
    var n = seguro(d.processo.numero);
    if (!n) return 'IPTU_Rascunho_' + seguro(String(d.id).slice(0, 8)) + '.json';
    var a = seguro(d.processo.ano);
    return 'IPTU_Processo_' + n + (a ? '-' + a : '') + '.json';
  }

  function nomeReplica(d) {
    var prot = seguro(T.protocolo(d).replace('/', '-')) || 'sem-numero';
    var dec = T.nomeDecisao(d.decisao.resultado) || 'SEM-DECISAO';
    return 'Replica_Fiscal_' + prot + '_' + dec + '.docx';
  }

  function aviso(msg, tipo, ms) {
    var el = document.createElement('div');
    el.className = 'aviso-toast ' + (tipo || '');
    el.innerHTML = msg;
    $('#avisos').appendChild(el);
    setTimeout(function () { el.remove(); }, ms || (tipo === 'erro' ? 8000 : 4500));
  }

  /** Diálogo com botões personalizados. Resolve com o "valor" do botão clicado (ou null). */
  function perguntar(titulo, html, botoes) {
    var dlg = $('#dlg-pergunta');
    $('#pergunta-titulo').textContent = titulo;
    $('#pergunta-texto').innerHTML = html;
    var caixa = $('#pergunta-botoes');
    caixa.innerHTML = '';
    return new Promise(function (ok) {
      function fechar(v) { dlg.removeEventListener('close', aoFechar); dlg.close(); ok(v); }
      function aoFechar() { dlg.removeEventListener('close', aoFechar); ok(null); }
      (botoes || [{ rotulo: 'OK', valor: true, estilo: 'primario' }]).forEach(function (b) {
        var bt = document.createElement('button');
        bt.type = 'button';
        bt.className = 'btn ' + (b.estilo || '');
        bt.textContent = b.rotulo;
        bt.addEventListener('click', function () { fechar(b.valor); });
        caixa.appendChild(bt);
      });
      dlg.addEventListener('close', aoFechar);
      dlg.showModal();
    });
  }

  function confirmar(titulo, html, rotuloOk, estiloOk) {
    return perguntar(titulo, html, [
      { rotulo: 'Cancelar', valor: false },
      { rotulo: rotuloOk || 'Confirmar', valor: true, estilo: estiloOk || 'primario' }
    ]).then(function (v) { return v === true; });
  }

  function erroAmigavel(e) {
    if (!e) return 'Erro desconhecido.';
    if (e.name === 'AbortError') return '';
    if (e.name === 'NotAllowedError' || e.name === 'SecurityError') return 'O navegador não autorizou o acesso à pasta.';
    if (e.name === 'NotFoundError') return 'Arquivo ou pasta não encontrado(a). A pasta pode ter sido movida ou renomeada.';
    if (e.name === 'NoModificationAllowedError' || e.name === 'InvalidModificationError') return 'O arquivo está bloqueado ou a pasta é somente leitura.';
    return e.message || String(e);
  }

  // ---------------------------------------------------------------- modelo de dados
  function processoNovo() {
    var ano = String(new Date().getFullYear());
    return {
      formato: FORMATO,
      versao: VERSAO,
      id: gerarId(),
      criadoEm: agoraISO(),
      atualizadoEm: agoraISO(),
      situacao: 'rascunho',
      processo: { numero: '', ano: ano, protocolo: '', dataEntrada: '', exercicio: ano, cci: '', enderecoImovel: '' },
      requerente: { nome: '', documento: '', contato: '', alegacoes: '' },
      analise: {
        tecnicoNome: '', dataRelatorio: '', opiniaoTecnico: '', constatacao: '',
        areaConfirmada: '', areaLancada: '', areaApurada: '', constatacaoTexto: '', areaAlegada: '',
        paragrafoComplementar: '', observacoes: ''
      },
      decisao: { resultado: '', justificativaDivergencia: '', data: hoje(), incluirData: false },
      fiscal: clonar(estado.config.fiscal),
      replica: { personalizada: false, texto: '', espacoReduzido: false, emissoes: [] }
    };
  }

  function normalizar(dados) {
    var base = processoNovo();
    base.fiscal = { nome: '', cargo: '', matricula: '' };
    var d = mesclar(base, clonar(dados));
    if (!Array.isArray(d.replica.emissoes)) d.replica.emissoes = [];
    return d;
  }

  function pareceProcesso(o) {
    return o && typeof o === 'object' && (o.formato === FORMATO || (o.processo && o.requerente));
  }

  // ---------------------------------------------------------------- configuração
  function carregarConfigLocal() {
    try {
      var c = JSON.parse(localStorage.getItem(CHAVE_LOCAL) || 'null');
      if (c && c.fiscal) estado.config.fiscal = mesclar(clonar(FISCAL_PADRAO), c.fiscal);
    } catch (e) { /* armazenamento indisponível */ }
  }

  async function carregarConfigPasta() {
    if (!P.suportado || !P.conectada()) return;
    try {
      if (!(await P.existe(ARQ_CONFIG))) return;
      var c = JSON.parse(await P.ler(ARQ_CONFIG));
      if (c && c.fiscal) estado.config.fiscal = mesclar(clonar(FISCAL_PADRAO), c.fiscal);
    } catch (e) { /* ignora configuração inválida */ }
  }

  async function salvarConfig() {
    try { localStorage.setItem(CHAVE_LOCAL, JSON.stringify(estado.config)); } catch (e) { /* ignora */ }
    if (P.suportado && P.conectada() && estado.pasta === 'conectada') {
      try {
        await P.gravar(ARQ_CONFIG, JSON.stringify({ formato: 'impugnacao-iptu/configuracao', fiscal: estado.config.fiscal }, null, 2));
      } catch (e) { aviso('Não foi possível gravar a configuração na pasta: ' + esc(erroAmigavel(e)), 'erro'); }
    }
  }

  // ---------------------------------------------------------------- pasta
  async function escolherPasta() {
    try {
      await P.escolher();
    } catch (e) {
      var m = erroAmigavel(e);
      if (m) aviso(esc(m), 'erro');
      return false;
    }
    estado.pasta = 'conectada';
    if (estado.atual) estado.arquivoAtual = null; // o processo aberto será gravado na nova pasta
    await carregarConfigPasta();
    await carregarLista();
    renderBarra();
    aviso('Pasta “' + esc(P.nome()) + '” conectada.', 'ok');
    return true;
  }

  async function reconectarPasta() {
    try {
      await P.reconectar();
    } catch (e) {
      aviso(esc(erroAmigavel(e)), 'erro');
      return false;
    }
    estado.pasta = 'conectada';
    await carregarConfigPasta();
    await carregarLista();
    renderBarra();
    return true;
  }

  /** Garante que há onde gravar. Retorna false se o usuário desistir. */
  async function garantirPasta() {
    if (!P.suportado) return true; // modo download
    if (estado.pasta === 'conectada') return true;
    if (estado.pasta === 'reconectar') return reconectarPasta();
    var ok = await confirmar('Escolha a pasta dos processos',
      '<p>Para salvar, indique a pasta onde os arquivos <code>.json</code> dos processos serão gravados.</p>',
      'Escolher pasta…');
    return ok ? escolherPasta() : false;
  }

  async function carregarLista() {
    var brutos = [];
    try {
      brutos = await P.listarJson();
    } catch (e) {
      aviso('Não foi possível ler a pasta: ' + esc(erroAmigavel(e)), 'erro');
      if (P.suportado) { estado.pasta = 'reconectar'; renderBarra(); }
    }
    estado.processos = [];
    estado.invalidos = [];
    brutos.forEach(function (b) {
      if (/^_/.test(b.arquivo)) return;
      if (b.erro) { estado.invalidos.push({ arquivo: b.arquivo, erro: b.erro }); return; }
      try {
        var o = JSON.parse(b.texto);
        if (!pareceProcesso(o)) throw new Error('não é um processo do assistente');
        estado.processos.push({ arquivo: b.arquivo, dados: normalizar(o), modificado: b.modificado });
      } catch (e) {
        estado.invalidos.push({ arquivo: b.arquivo, erro: e.message });
      }
    });
    renderLista();
  }

  // ---------------------------------------------------------------- barra e boas-vindas
  function renderBarra() {
    var chip = $('#btn-pasta');
    var nome = $('#nome-pasta');
    chip.classList.remove('ok', 'atencao');
    if (estado.pasta === 'conectada') {
      nome.textContent = P.nome();
      chip.classList.add('ok');
      chip.title = 'Pasta dos processos: ' + P.nome() + ' — clique para trocar';
    } else if (estado.pasta === 'reconectar') {
      nome.textContent = P.nome() + ' (reconectar)';
      chip.classList.add('atencao');
      chip.title = 'Clique para reconectar à pasta';
    } else if (estado.pasta === 'download') {
      nome.textContent = estado.nomePastaDownload ? estado.nomePastaDownload + ' (somente leitura)' : 'Modo download';
      chip.classList.add('atencao');
      chip.title = 'Este navegador não grava direto na pasta — os arquivos são baixados';
    } else {
      nome.textContent = 'Escolher pasta';
      chip.title = 'Escolher a pasta dos processos';
    }

    var bv = $('#boas-vindas');
    $('#bv-nenhuma').hidden = estado.pasta !== 'nenhuma';
    $('#bv-reconectar').hidden = estado.pasta !== 'reconectar';
    $('#bv-download').hidden = estado.pasta !== 'download';
    $('#bv-nome-pasta').textContent = P.nome();
    bv.hidden = estado.pasta === 'conectada';
    $('#painel-lista').hidden = estado.pasta === 'nenhuma' || estado.pasta === 'reconectar';

    var resumo = $('#resumo-pasta');
    if (estado.pasta === 'conectada') {
      resumo.innerHTML = 'Pasta: <b>' + esc(P.nome()) + '</b> · réplicas em <b>' + esc(P.nome()) + '/' + P.SUBPASTA_REPLICAS + '</b>';
    } else if (estado.pasta === 'download') {
      resumo.textContent = estado.nomePastaDownload
        ? 'Processos carregados da pasta “' + estado.nomePastaDownload + '”. As gravações geram downloads.'
        : 'Nenhuma pasta carregada. As gravações geram downloads.';
    } else {
      resumo.textContent = '';
    }
  }

  async function acaoChipPasta() {
    if (estado.pasta === 'download') { $('#entrada-pasta').click(); return; }
    if (estado.pasta === 'reconectar') { await reconectarPasta(); return; }
    if (estado.pasta === 'conectada' && estado.atual && estado.sujo) {
      var ok = await confirmar('Trocar de pasta',
        '<p>O processo aberto tem alterações não salvas. Ao trocar de pasta, ele será gravado na nova pasta quando você salvar.</p>',
        'Trocar pasta');
      if (!ok) return;
    }
    await escolherPasta();
  }

  // ---------------------------------------------------------------- lista
  function filtrados() {
    var q = semAcento($('#busca').value.trim().toLowerCase());
    var fd = $('#filtro-decisao').value;
    var fs = $('#filtro-situacao').value;
    return estado.processos.filter(function (p) {
      var d = p.dados;
      if (fd === 'sem' && d.decisao.resultado) return false;
      if (fd && fd !== 'sem' && d.decisao.resultado !== fd) return false;
      if (fs && d.situacao !== fs) return false;
      if (!q) return true;
      var alvo = semAcento([T.protocolo(d), d.processo.cci, d.requerente.nome, d.requerente.documento, p.arquivo].join(' ').toLowerCase());
      return alvo.indexOf(q) >= 0;
    }).sort(function (a, b) {
      return String(b.dados.atualizadoEm).localeCompare(String(a.dados.atualizadoEm));
    });
  }

  function etiquetaDecisao(r) {
    if (r === 'deferimento') return '<span class="etiqueta deferimento">Deferimento</span>';
    if (r === 'indeferimento') return '<span class="etiqueta indeferimento">Indeferimento</span>';
    return '<span class="etiqueta neutra">Sem decisão</span>';
  }

  function etiquetaSituacao(s) {
    return s === 'emitida'
      ? '<span class="etiqueta emitida">Réplica emitida</span>'
      : '<span class="etiqueta rascunho">Rascunho</span>';
  }

  function renderLista() {
    var todos = estado.processos.map(function (p) { return p.dados; });
    var cont = function (f) { return todos.filter(f).length; };
    var ind = [
      { v: todos.length, r: 'Processos' },
      { v: cont(function (d) { return d.situacao !== 'emitida'; }), r: 'Rascunhos' },
      { v: cont(function (d) { return d.situacao === 'emitida'; }), r: 'Réplicas emitidas' },
      { v: cont(function (d) { return d.decisao.resultado === 'deferimento'; }), r: 'Deferimentos', c: 'deferido' },
      { v: cont(function (d) { return d.decisao.resultado === 'indeferimento'; }), r: 'Indeferimentos', c: 'indeferido' }
    ];
    $('#indicadores').innerHTML = ind.map(function (i) {
      return '<div class="indicador ' + (i.c || '') + '"><div class="valor">' + i.v + '</div><div class="rotulo">' + i.r + '</div></div>';
    }).join('');

    var lista = filtrados();
    $('#corpo-tabela').innerHTML = lista.map(function (p) {
      var d = p.dados;
      var prot = T.protocolo(d);
      return '<tr data-arquivo="' + esc(p.arquivo) + '">' +
        '<td class="num forte">' + (prot ? esc(prot) : '<span class="etiqueta neutra">sem número</span>') + '</td>' +
        '<td>' + esc(d.requerente.nome || '—') + '</td>' +
        '<td class="num">' + esc(d.processo.cci || '—') + '</td>' +
        '<td class="num">' + esc(d.processo.exercicio || '—') + '</td>' +
        '<td>' + etiquetaDecisao(d.decisao.resultado) + '</td>' +
        '<td>' + etiquetaSituacao(d.situacao) + '</td>' +
        '<td class="num">' + esc(dataHora(d.atualizadoEm)) + '</td>' +
        '<td class="col-acoes"><button type="button" class="btn pequeno fantasma perigo" data-excluir="' + esc(p.arquivo) + '" title="Excluir o arquivo do processo">Excluir</button></td>' +
        '</tr>';
    }).join('');

    var vazio = $('#lista-vazia');
    if (!estado.processos.length) {
      vazio.hidden = false;
      vazio.innerHTML = '<p>Nenhum processo nesta pasta ainda.</p><div class="acoes"><button type="button" class="btn primario" data-acao="novo">+ Novo processo</button></div>';
    } else if (!lista.length) {
      vazio.hidden = false;
      vazio.innerHTML = '<p>Nenhum processo corresponde aos filtros.</p>';
    } else {
      vazio.hidden = true;
    }

    var inv = $('#aviso-invalidos');
    inv.hidden = !estado.invalidos.length;
    if (estado.invalidos.length) {
      inv.textContent = estado.invalidos.length + ' arquivo(s) .json da pasta foram ignorados por não serem processos válidos: ' +
        estado.invalidos.map(function (i) { return i.arquivo; }).join(', ') + '.';
    }
  }

  function exportarCsv() {
    var cab = ['Processo', 'Data de entrada', 'Exercício', 'CCI', 'Impugnante', 'CPF/CNPJ', 'Opinião técnica',
      'Decisão', 'Situação', 'Última emissão', 'Atualizado em', 'Arquivo'];
    var linhas = filtrados().map(function (p) {
      var d = p.dados;
      var ult = d.replica.emissoes[d.replica.emissoes.length - 1];
      return [T.protocolo(d), T.dataCurta(d.processo.dataEntrada), d.processo.exercicio, d.processo.cci,
        d.requerente.nome, d.requerente.documento, d.analise.opiniaoTecnico, d.decisao.resultado,
        d.situacao === 'emitida' ? 'Réplica emitida' : 'Rascunho', ult ? dataHora(ult.em) : '',
        dataHora(d.atualizadoEm), p.arquivo];
    });
    var csv = [cab].concat(linhas).map(function (l) {
      return l.map(function (c) {
        c = String(c === undefined || c === null ? '' : c);
        return /[";\n\r]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c;
      }).join(';');
    }).join('\r\n');
    P.baixar('Processos_Impugnacao_IPTU_' + hoje() + '.csv', new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' }));
  }

  async function excluir(arquivo) {
    var item = estado.processos.find(function (p) { return p.arquivo === arquivo; });
    if (!item) return;
    var prot = T.protocolo(item.dados) || 'sem número';
    var html = estado.pasta === 'conectada'
      ? '<p>Excluir o processo <b>' + esc(prot) + '</b>' + (item.dados.requerente.nome ? ' (' + esc(item.dados.requerente.nome) + ')' : '') +
        '?</p><p>O arquivo <code>' + esc(arquivo) + '</code> será apagado da pasta. Esta ação não pode ser desfeita. ' +
        'Réplicas .docx já geradas não são apagadas.</p>'
      : '<p>Remover o processo <b>' + esc(prot) + '</b> da lista? No modo download o arquivo original não é apagado.</p>';
    if (!(await confirmar('Excluir processo', html, 'Excluir', 'perigo'))) return;
    try {
      await P.remover(arquivo);
      estado.processos = estado.processos.filter(function (p) { return p.arquivo !== arquivo; });
      renderLista();
      aviso('Processo ' + esc(prot) + ' excluído.', 'ok');
    } catch (e) {
      aviso('Não foi possível excluir: ' + esc(erroAmigavel(e)), 'erro');
    }
  }

  // ---------------------------------------------------------------- editor
  function abrirEditor(dados, arquivo) {
    estado.atual = normalizar(dados);
    estado.arquivoAtual = arquivo || null;
    estado.sujo = false;
    estado.salvoEm = arquivo ? estado.atual.atualizadoEm : null;
    estado.etapa = 0;
    preencherFormulario();
    $('#tela-lista').hidden = true;
    $('#tela-editor').hidden = false;
    irPara(0);
    atualizarTudo();
    window.scrollTo(0, 0);
  }

  async function fecharEditor() {
    if (estado.sujo) {
      var r = await perguntar('Alterações não salvas',
        '<p>Este processo tem alterações que ainda não foram salvas.</p>', [
          { rotulo: 'Cancelar', valor: 'cancelar' },
          { rotulo: 'Sair sem salvar', valor: 'sair', estilo: 'perigo' },
          { rotulo: 'Salvar e sair', valor: 'salvar', estilo: 'primario' }
        ]);
      if (r === 'salvar') { if (!(await salvar())) return; }
      else if (r !== 'sair') return;
    }
    estado.atual = null;
    estado.arquivoAtual = null;
    estado.sujo = false;
    $('#tela-editor').hidden = true;
    $('#tela-lista').hidden = false;
    renderLista();
  }

  function campos() { return $$('#formulario [name]'); }

  function preencherFormulario() {
    var d = estado.atual;
    campos().forEach(function (el) {
      var v = obter(d, el.name);
      if (el.type === 'radio') el.checked = v === el.value;
      else if (el.type === 'checkbox') el.checked = !!v;
      else el.value = v === undefined || v === null ? '' : v;
      el.removeAttribute('aria-invalid');
    });
  }

  function aoEditar(ev) {
    var el = ev.target;
    if (!el.name || !estado.atual) return;
    var d = estado.atual;
    var valor = el.type === 'checkbox' ? el.checked : el.value;
    if (el.type === 'radio' && !el.checked) return;
    definir(d, el.name, valor);
    el.removeAttribute('aria-invalid');

    if (el.name === 'analise.opiniaoTecnico' && !d.analise.constatacao) {
      // Pré-seleção usual (o(a) fiscal pode alterar): indeferimento → área confirmada; deferimento → área retificada.
      d.analise.constatacao = valor === 'indeferimento' ? 'confirmada' : 'retificada';
      $$('[name="analise.constatacao"]').forEach(function (r) { r.checked = r.value === d.analise.constatacao; });
    }
    if (el.name === 'replica.personalizada' && valor) {
      if (!T.limpo(d.replica.texto)) d.replica.texto = T.textoAutomatico(d);
      $('[name="replica.texto"]').value = d.replica.texto;
    }
    marcarSujo();
    atualizarTudo();
  }

  function marcarSujo() {
    estado.sujo = true;
    renderEstadoSalvo();
  }

  function renderEstadoSalvo() {
    var el = $('#estado-salvo');
    if (estado.sujo) {
      el.textContent = '● Alterações não salvas';
      el.className = 'estado-salvo sujo';
    } else if (!estado.arquivoAtual) {
      el.textContent = 'Ainda não salvo';
      el.className = 'estado-salvo';
    } else {
      el.textContent = 'Salvo em ' + dataHora(estado.salvoEm) + ' · ' + estado.arquivoAtual;
      el.className = 'estado-salvo';
    }
  }

  var temporizadorPrevia = null;
  function atualizarTudo() {
    var d = estado.atual;
    if (!d) return;
    var prot = T.protocolo(d);
    $('#titulo-processo').textContent = prot ? 'Processo nº ' + prot : 'Novo processo';
    renderEstadoSalvo();
    atualizarCondicionais();
    atualizarEtapas();
    if (estado.etapa === 3) renderRevisao();
    clearTimeout(temporizadorPrevia);
    temporizadorPrevia = setTimeout(renderPrevia, 80);
  }

  function atualizarCondicionais() {
    var d = estado.atual;
    $$('[data-se]').forEach(function (el) {
      var regra = el.getAttribute('data-se').split('=');
      el.hidden = obter(d, regra[0]) !== regra[1];
    });

    // diferença de área (retificação)
    var dif = $('#diferenca-area');
    var la = numero(d.analise.areaLancada), ap = numero(d.analise.areaApurada);
    if (!isNaN(la) && !isNaN(ap) && la > 0) {
      var delta = ap - la;
      var pct = delta / la * 100;
      dif.textContent = 'Diferença: ' + (delta > 0 ? '+' : '') + delta.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) +
        ' m² (' + (pct > 0 ? '+' : '') + pct.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + '%) em relação à área lançada.';
    } else {
      dif.textContent = '';
    }

    // decisão
    var op = d.analise.opiniaoTecnico;
    $('#ref-opiniao').innerHTML = op
      ? 'Opinião registrada no relatório técnico: ' + etiquetaDecisao(op)
      : 'Opinião do relatório técnico: <i>não informada</i> (etapa 2).';
    $('#bloco-divergencia').hidden = !T.haDivergencia(d);

    $('#bloco-personalizado').hidden = !d.replica.personalizada;
  }

  function numero(v) {
    var s = T.limpo(v).replace(/m²|m2/gi, '').replace(/\s/g, '');
    if (!s) return NaN;
    if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
    return /^\d+(\.\d+)?$/.test(s) ? Number(s) : NaN;
  }

  function atualizarEtapas() {
    var pend = T.pendencias(estado.atual);
    $$('#etapas button').forEach(function (b) {
      var i = Number(b.getAttribute('data-ir'));
      b.classList.toggle('ativa', i === estado.etapa);
      var completa = i < 3
        ? !pend.some(function (p) { return p.etapa === i; })
        : !pend.length && estado.atual.situacao === 'emitida';
      b.classList.toggle('completa', completa);
      b.setAttribute('aria-current', i === estado.etapa ? 'step' : 'false');
    });
  }

  function irPara(i, focarCampo) {
    estado.etapa = Math.max(0, Math.min(3, i));
    $$('#formulario .etapa').forEach(function (s) { s.hidden = Number(s.getAttribute('data-etapa')) !== estado.etapa; });
    $('#btn-anterior').style.visibility = estado.etapa === 0 ? 'hidden' : 'visible';
    $('#btn-proxima').hidden = estado.etapa === 3;
    atualizarEtapas();
    if (estado.etapa === 3) renderRevisao();
    if (focarCampo) {
      var el = $('#formulario [name="' + focarCampo + '"]');
      if (el) {
        el.setAttribute('aria-invalid', 'true');
        setTimeout(function () { el.focus(); el.scrollIntoView({ block: 'center', behavior: 'smooth' }); }, 30);
      }
    } else {
      window.scrollTo(0, 0);
    }
  }

  function renderRevisao() {
    var d = estado.atual;
    var a = d.analise;
    var constat = { confirmada: 'Área construída confirmada', retificada: 'Área construída retificada', outra: 'Outra constatação' }[a.constatacao] || '—';
    var linhas = [
      ['Processo', T.protocolo(d) || '—'],
      ['Impugnante', d.requerente.nome || '—'],
      ['CCI / exercício', (d.processo.cci || '—') + ' / ' + (d.processo.exercicio || '—')],
      ['Relatório técnico', (a.opiniaoTecnico ? 'pelo ' + a.opiniaoTecnico : 'opinião não informada') + ' · ' + constat],
      ['Decisão do(a) fiscal', d.decisao.resultado ? etiquetaDecisao(d.decisao.resultado) : '<span class="etiqueta neutra">não definida</span>'],
      ['Assinatura', esc(T.limpo(d.fiscal.cargo) || '—') + ' – Matrícula nº ' + esc(T.limpo(d.fiscal.matricula) || '—')]
    ];
    $('#resumo-revisao').innerHTML = '<dl class="resumo">' + linhas.map(function (l, i) {
      var v = (i === 4 || i === 5) ? l[1] : esc(l[1]);
      return '<dt>' + l[0] + '</dt><dd>' + v + '</dd>';
    }).join('') + '</dl>';

    var pend = T.pendencias(d);
    var nomesEtapa = ['Processo e impugnante', 'Relatório técnico', 'Decisão', 'Revisão'];
    $('#pendencias').innerHTML = pend.length
      ? '<div class="pendencias"><b>Antes de emitir, complete:</b><ul>' + pend.map(function (p) {
          return '<li>' + esc(p.texto) + ' <small>(' + nomesEtapa[p.etapa] + ')</small>' +
            '<button type="button" data-corrigir="' + p.etapa + '|' + esc(p.campo) + '">corrigir</button></li>';
        }).join('') + '</ul></div>'
      : '<div class="pronto">✓ Todos os campos obrigatórios foram preenchidos. Confira a prévia antes de emitir.</div>';

    var destino = $('#destino-replica');
    destino.innerHTML = estado.pasta === 'conectada'
      ? 'A réplica será gravada em <code>' + esc(P.nome()) + '/' + P.SUBPASTA_REPLICAS + '/' + esc(nomeReplica(d)) + '</code> e o processo será salvo.'
      : P.suportado
        ? 'Ao emitir, você escolherá a pasta dos processos (se ainda não escolheu).'
        : 'Neste navegador, a réplica e o processo (.json) serão baixados.';

    var em = d.replica.emissoes || [];
    $('#historico').innerHTML = em.length
      ? '<div class="historico"><h3>Emissões anteriores</h3><ol>' + em.map(function (e) {
          return '<li>' + esc(dataHora(e.em)) + ' — ' + esc(T.nomeDecisao(e.resultado) || 'sem decisão') + ' — <code>' + esc(e.arquivo) + '</code></li>';
        }).join('') + '</ol></div>'
      : '';

    var ultimo = $('#btn-baixar-copia');
    if (ultimo) ultimo.remove();
    if (estado.ultimaReplica && estado.ultimaReplica.id === d.id && estado.pasta === 'conectada') {
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'btn pequeno';
      b.id = 'btn-baixar-copia';
      b.textContent = 'Baixar uma cópia da última réplica';
      b.addEventListener('click', function () { P.baixar(estado.ultimaReplica.nome, estado.ultimaReplica.blob); });
      destino.appendChild(document.createElement('br'));
      destino.appendChild(b);
    }
  }

  // ---------------------------------------------------------------- documento (prévia / impressão)
  function htmlTexto(texto, impressao) {
    return esc(texto)
      .replace(/⟦([^⟧]*)⟧/g, function (t, rot) { return impressao ? '_____' : '<mark class="pendente">' + rot + '</mark>'; })
      .replace(/\n/g, '<br>');
  }

  function htmlParagrafo(p, impressao) {
    return T.trechos(p).map(function (tr) {
      var h = htmlTexto(tr.texto, impressao);
      return tr.negrito ? '<b>' + h + '</b>' : h;
    }).join('');
  }

  function htmlFolha(doc, impressao) {
    var cab = '<table class="cab-doc"><tr>' +
      '<td class="c-logo"><img src="' + estado.imagens.brasao + '" alt="Brasão do Município de Araguari"></td>' +
      '<td class="c-texto"><b>PREFEITURA MUNICIPAL DE ARAGUARI / MG</b><b>SECRETARIA MUNICIPAL DE FAZENDA – SEFAZ</b>' +
      '<b>DEPARTAMENTO DE FISCALIZAÇÃO DE RENDAS/ISS</b><span>Avenida Senador Melo Viana, 176, B. Goiás – CEP 38442-192</span>' +
      '<span>(34) 3690-3013 / fiscaliss@araguari.mg.gov.br</span></td>' +
      '<td class="c-num">PROCESSO<br>Nº ' + esc(doc.numeroCabecalho) + '</td>' +
      '</tr></table>';
    var corpo = doc.paragrafos.map(function (p) { return '<p>' + htmlParagrafo(p, impressao) + '</p>'; }).join('');
    if (doc.localData) corpo += '<p class="local-data">' + htmlParagrafo(doc.localData, impressao) + '</p>';
    var ass = (doc.assinaturaNome ? '<p>' + esc(doc.assinaturaNome) + '</p>' : '') +
      '<p>' + htmlTexto(doc.assinaturaCargo, impressao) + '</p>';
    return '<div class="folha' + (impressao ? ' impressa' : '') + '">' +
      '<img class="marca-dagua" src="' + estado.imagens.marca + '" alt="">' +
      '<table class="estrutura"><thead><tr><td>' + cab + '</td></tr></thead>' +
      '<tbody><tr><td><div class="corpo-doc">' + corpo + '</div>' +
      '<div class="assinatura-doc' + (doc.espacoReduzido ? ' reduzido' : '') + '">' + ass + '</div></td></tr></tbody></table>' +
      '</div>';
  }

  function renderPrevia() {
    var d = estado.atual;
    if (!d) return;
    var doc = T.montarDocumento(d);
    var alvo = $('#previa');
    alvo.innerHTML = htmlFolha(doc, false);
    ajustarZoom();

    var selo = $('#selo-previa');
    selo.innerHTML = d.decisao.resultado ? etiquetaDecisao(d.decisao.resultado) : '';

    var avisos = [];
    if (d.replica.personalizada) avisos.push('Texto editado manualmente — mudanças nos campos não alteram o corpo da réplica.');
    var folha = $('.folha', alvo);
    // A folha tem altura mínima de uma página A4; se crescer além disso, o texto não cabe em uma página.
    // Tolerância de meia linha: a prévia usa as fontes do navegador e pode diferir levemente do Word.
    var excesso = folha ? folha.offsetHeight - folha.offsetWidth * 297 / 210 : 0;
    if (excesso > 16) {
      avisos.push('A réplica deve ocupar mais de uma página.' +
        (d.replica.espacoReduzido ? '' : ' Se preferir, marque “Reduzir o espaço antes da assinatura” na etapa 4.'));
    }
    var av = $('#aviso-previa');
    av.hidden = !avisos.length;
    av.textContent = avisos.join(' ');
  }

  var pxPorMm = null;
  function mmParaPx(mm) {
    if (pxPorMm === null) {
      var t = document.createElement('div');
      t.style.cssText = 'position:absolute;visibility:hidden;width:100mm;height:0';
      document.body.appendChild(t);
      pxPorMm = t.getBoundingClientRect().width / 100;
      t.remove();
    }
    return mm * pxPorMm;
  }

  function ajustarZoom() {
    var area = $('#previa-area');
    var folha = $('#previa .folha');
    if (!area || !folha) return;
    var largura = area.clientWidth - 12;
    var z = Math.min(1, largura / mmParaPx(210));
    folha.style.zoom = z > 0 ? String(z) : '1';
  }

  // ---------------------------------------------------------------- salvar / emitir
  async function salvar(opcoes) {
    opcoes = opcoes || {};
    var d = estado.atual;
    if (!d) return false;
    if (!(await garantirPasta())) return false;

    d.processo.protocolo = T.protocolo(d);
    d.formato = FORMATO;
    d.versao = VERSAO;
    var nome = nomeArquivo(d);

    try {
      if (nome !== estado.arquivoAtual && (await P.existe(nome))) {
        var outro = null;
        try { outro = JSON.parse(await P.ler(nome)); } catch (e) { /* ilegível */ }
        if (!outro || outro.id !== d.id) {
          await perguntar('Número de processo já existente',
            '<p>Já existe na pasta o arquivo <code>' + esc(nome) + '</code>, de outro processo' +
            (outro && outro.requerente && outro.requerente.nome ? ' (impugnante: <b>' + esc(outro.requerente.nome) + '</b>)' : '') +
            '.</p><p>Confira o número e o ano do processo. Nada foi gravado.</p>');
          irPara(0, 'processo.numero');
          return false;
        }
      }
      d.atualizadoEm = agoraISO();
      var modo = await P.gravar(nome, JSON.stringify(d, null, 2));
      if (estado.arquivoAtual && estado.arquivoAtual !== nome) {
        try { await P.remover(estado.arquivoAtual); } catch (e) { /* arquivo antigo já não existe */ }
      }
      var antigo = estado.arquivoAtual;
      estado.processos = estado.processos.filter(function (p) { return p.arquivo !== nome && p.arquivo !== antigo; });
      estado.processos.push({ arquivo: nome, dados: clonar(d), modificado: Date.now() });
      estado.arquivoAtual = nome;
      estado.sujo = false;
      estado.salvoEm = d.atualizadoEm;
      renderEstadoSalvo();
      atualizarEtapas();
      if (!opcoes.silencioso) {
        aviso(modo === 'pasta'
          ? 'Processo salvo em <b>' + esc(P.nome() + '/' + nome) + '</b>.'
          : 'Arquivo <b>' + esc(nome) + '</b> baixado. Guarde-o na pasta dos processos.', 'ok');
      }
      return true;
    } catch (e) {
      aviso('Não foi possível salvar: ' + esc(erroAmigavel(e)), 'erro');
      return false;
    }
  }

  async function emitirDocx() {
    var d = estado.atual;
    var pend = T.pendencias(d);
    if (pend.length) {
      var ok = await confirmar('Há campos pendentes',
        '<p>Os itens abaixo ainda não foram preenchidos e sairão como “_____” no documento:</p><ul>' +
        pend.map(function (p) { return '<li>' + esc(p.texto) + '</li>'; }).join('') + '</ul>',
        'Gerar mesmo assim');
      if (!ok) return;
    }
    if (!(await garantirPasta())) return;

    var doc = T.montarDocumento(d);
    var prot = T.protocolo(d);
    var nome = nomeReplica(d);
    var blob;
    try {
      blob = D.gerar(doc, 'Réplica Fiscal – Processo nº ' + (prot || 's/n'),
        T.limpo(d.fiscal.nome) || T.limpo(d.fiscal.cargo) || 'Departamento de Fiscalização de Rendas/ISS');
    } catch (e) {
      aviso('Falha ao montar o documento: ' + esc(e.message), 'erro');
      return;
    }

    var modo;
    try {
      modo = await P.gravarReplica(nome, blob);
    } catch (e) {
      aviso('Não foi possível gravar a réplica: ' + esc(erroAmigavel(e)), 'erro');
      return;
    }

    d.replica.emissoes.push({
      em: agoraISO(),
      resultado: d.decisao.resultado,
      arquivo: nome,
      fiscal: clonar(d.fiscal),
      texto: doc.paragrafos.map(T.paraTextoSimples)
    });
    d.situacao = 'emitida';
    estado.ultimaReplica = { id: d.id, nome: nome, blob: blob };
    marcarSujo();
    var salvo = await salvar({ silencioso: true });
    renderRevisao();
    atualizarEtapas();

    if (modo === 'pasta') {
      aviso('Réplica gravada em <b>' + esc(P.nome() + '/' + P.SUBPASTA_REPLICAS + '/' + nome) + '</b>' +
        (salvo ? ' e processo salvo.' : '. <b>Atenção:</b> o processo não foi salvo.'), salvo ? 'ok' : 'erro', 7000);
    } else {
      aviso('Réplica <b>' + esc(nome) + '</b> baixada' + (salvo ? ', junto com o arquivo do processo.' : '.'), 'ok', 7000);
    }
  }

  // Área útil da página A4 na impressão (margens do modelo: sup. 1,25 cm até o cabeçalho, inf. 2 cm).
  var ALTURA_UTIL_MM = 297 - 12.5 - 20;

  function imprimir() {
    var d = estado.atual;
    var area = $('#impressao');
    area.innerHTML = htmlFolha(T.montarDocumento(d), true);

    // Se o texto passar de uma página por poucos milímetros, encurta o espaço antes da assinatura
    // para não gerar uma segunda página só com a assinatura (mantém ao menos 1 cm).
    area.classList.add('medindo');
    var folha = $('.folha', area), ass = $('.assinatura-doc', area);
    var excesso = folha.offsetHeight - mmParaPx(ALTURA_UTIL_MM) + mmParaPx(1.5);
    var espaco = parseFloat(getComputedStyle(ass).marginTop) || 0;
    if (excesso > 0 && espaco - excesso >= mmParaPx(10)) ass.style.marginTop = (espaco - excesso) + 'px';
    area.classList.remove('medindo');
    var titulo = document.title;
    document.title = nomeReplica(d).replace(/\.docx$/, '');
    var restaurar = function () { document.title = titulo; window.removeEventListener('afterprint', restaurar); };
    window.addEventListener('afterprint', restaurar);
    setTimeout(function () { window.print(); }, 50);
  }

  async function copiarTexto() {
    var doc = T.montarDocumento(estado.atual);
    var partes = doc.paragrafos.slice();
    if (doc.localData) partes.push(doc.localData);
    partes.push((doc.assinaturaNome ? doc.assinaturaNome + '\n' : '') + doc.assinaturaCargo);
    var texto = partes.map(T.paraTextoSimples).join('\n\n');
    try {
      await navigator.clipboard.writeText(texto);
    } catch (e) {
      var ta = document.createElement('textarea');
      ta.value = texto;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      ta.remove();
    }
    aviso('Texto da réplica copiado.', 'ok');
  }

  // ---------------------------------------------------------------- configuração (diálogo)
  function abrirConfig() {
    var f = $('#form-config');
    f.nome.value = estado.config.fiscal.nome || '';
    f.cargo.value = estado.config.fiscal.cargo || '';
    f.matricula.value = estado.config.fiscal.matricula || '';
    $('#dlg-config').showModal();
  }

  async function salvarDialogConfig(ev) {
    ev.preventDefault();
    var f = $('#form-config');
    estado.config.fiscal = { nome: f.nome.value.trim(), cargo: f.cargo.value.trim(), matricula: f.matricula.value.trim() };
    $('#dlg-config').close();
    await salvarConfig();
    aviso('Dados do(a) fiscal atualizados. Eles serão usados nos novos processos.', 'ok');
  }

  // ---------------------------------------------------------------- inicialização
  function criarImagens() {
    function url(chave) {
      var bin = atob(M.imagens[chave]);
      var bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return URL.createObjectURL(new Blob([bytes], { type: 'image/png' }));
    }
    estado.imagens.brasao = url('word/media/image2.png');
    estado.imagens.marca = url('word/media/image1.png');
    $('#brasao-barra').src = estado.imagens.brasao;
  }

  function ligarEventos() {
    document.addEventListener('click', function (ev) {
      var alvo = ev.target.closest('[data-acao]');
      if (!alvo) return;
      var a = alvo.getAttribute('data-acao');
      if (a === 'escolher-pasta') escolherPasta();
      else if (a === 'reconectar') reconectarPasta();
      else if (a === 'carregar-pasta') $('#entrada-pasta').click();
      else if (a === 'novo') abrirEditor(processoNovo(), null);
    });

    $('#btn-pasta').addEventListener('click', acaoChipPasta);
    $('#btn-config').addEventListener('click', abrirConfig);
    $('#form-config').addEventListener('submit', salvarDialogConfig);
    $$('[data-fechar]').forEach(function (b) { b.addEventListener('click', function () { b.closest('dialog').close(); }); });

    $('#btn-novo').addEventListener('click', function () { abrirEditor(processoNovo(), null); });
    $('#btn-atualizar').addEventListener('click', carregarLista);
    $('#btn-csv').addEventListener('click', exportarCsv);
    ['#busca', '#filtro-decisao', '#filtro-situacao'].forEach(function (s) { $(s).addEventListener('input', renderLista); });

    $('#corpo-tabela').addEventListener('click', function (ev) {
      var del = ev.target.closest('[data-excluir]');
      if (del) { ev.stopPropagation(); excluir(del.getAttribute('data-excluir')); return; }
      var tr = ev.target.closest('tr[data-arquivo]');
      if (!tr) return;
      var item = estado.processos.find(function (p) { return p.arquivo === tr.getAttribute('data-arquivo'); });
      if (item) abrirEditor(item.dados, item.arquivo);
    });

    $('#entrada-pasta').addEventListener('change', async function (ev) {
      var arquivos = ev.target.files;
      if (!arquivos || !arquivos.length) return;
      estado.nomePastaDownload = await P.carregarArquivos(arquivos);
      ev.target.value = '';
      await carregarLista();
      renderBarra();
      aviso(estado.processos.length + ' processo(s) carregado(s).', 'ok');
    });

    var form = $('#formulario');
    function marcavel(el) { return el.type === 'radio' || el.type === 'checkbox'; }
    form.addEventListener('input', function (ev) { if (!marcavel(ev.target)) aoEditar(ev); });
    form.addEventListener('change', function (ev) { if (marcavel(ev.target)) aoEditar(ev); });
    form.addEventListener('submit', function (ev) { ev.preventDefault(); });

    $('#btn-voltar').addEventListener('click', fecharEditor);
    $('#btn-salvar').addEventListener('click', function () { salvar(); });
    $('#btn-anterior').addEventListener('click', function () { irPara(estado.etapa - 1); });
    $('#btn-proxima').addEventListener('click', function () { irPara(estado.etapa + 1); });
    $('#etapas').addEventListener('click', function (ev) {
      var b = ev.target.closest('[data-ir]');
      if (b) irPara(Number(b.getAttribute('data-ir')));
    });
    $('#pendencias').addEventListener('click', function (ev) {
      var b = ev.target.closest('[data-corrigir]');
      if (!b) return;
      var partes = b.getAttribute('data-corrigir').split('|');
      irPara(Number(partes[0]), partes[1]);
    });

    var sug = $('#sugestoes');
    SUGESTOES.forEach(function (s) {
      var b = document.createElement('button');
      b.type = 'button';
      b.textContent = s.rotulo;
      b.title = s.texto;
      b.addEventListener('click', function () {
        var ta = $('[name="analise.paragrafoComplementar"]');
        ta.value = T.limpo(ta.value) ? ta.value.replace(/\s+$/, '') + '\n\n' + s.texto : s.texto;
        ta.dispatchEvent(new Event('input', { bubbles: true }));
        ta.focus();
      });
      sug.appendChild(b);
    });

    $('#btn-usar-padrao').addEventListener('click', function () {
      estado.atual.fiscal = clonar(estado.config.fiscal);
      preencherFormulario();
      marcarSujo();
      atualizarTudo();
    });
    $('#btn-tornar-padrao').addEventListener('click', async function () {
      estado.config.fiscal = clonar(estado.atual.fiscal);
      await salvarConfig();
      aviso('Dados de assinatura salvos como padrão para os novos processos.', 'ok');
    });
    $('#btn-restaurar-texto').addEventListener('click', async function () {
      var ok = await confirmar('Restaurar texto automático',
        '<p>O texto editado manualmente será substituído pelo texto montado a partir dos campos.</p>', 'Restaurar');
      if (!ok) return;
      var d = estado.atual;
      d.replica.texto = T.textoAutomatico(d);
      $('[name="replica.texto"]').value = d.replica.texto;
      marcarSujo();
      atualizarTudo();
    });

    $('#btn-docx').addEventListener('click', emitirDocx);
    $('#btn-imprimir').addEventListener('click', imprimir);
    $('#btn-copiar').addEventListener('click', copiarTexto);

    document.addEventListener('keydown', function (ev) {
      if ((ev.ctrlKey || ev.metaKey) && (ev.key === 's' || ev.key === 'S')) {
        if (!$('#tela-editor').hidden) { ev.preventDefault(); salvar(); }
      }
    });
    window.addEventListener('beforeunload', function (ev) {
      if (estado.atual && estado.sujo) { ev.preventDefault(); ev.returnValue = ''; }
    });
    if (window.ResizeObserver) new ResizeObserver(ajustarZoom).observe($('#previa-area'));
    else window.addEventListener('resize', ajustarZoom);
  }

  async function iniciar() {
    carregarConfigLocal();
    criarImagens();
    ligarEventos();
    if (!P.suportado) {
      estado.pasta = 'download';
    } else {
      var r = await P.restaurar();
      if (r && r.permissao === 'granted') {
        estado.pasta = 'conectada';
        await carregarConfigPasta();
        await carregarLista();
      } else if (r) {
        estado.pasta = 'reconectar';
      }
    }
    renderBarra();
    renderLista();
  }

  iniciar();
})();
