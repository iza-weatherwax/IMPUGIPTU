/*
 * Armazenamento dos processos (.json) na pasta indicada pelo(a) fiscal.
 *
 * Modo "pasta" (Chrome/Edge): usa a File System Access API para ler e gravar diretamente na pasta.
 *   A pasta escolhida é lembrada entre usos (IndexedDB do próprio navegador).
 * Modo "download" (demais navegadores): os processos são carregados a partir de uma pasta
 *   selecionada e cada gravação gera um download do arquivo .json.
 */
(function (global) {
  'use strict';

  var BANCO = 'assistente-impugnacao-iptu';
  var LOJA = 'preferencias';
  var SUBPASTA_REPLICAS = 'replicas';

  var suportado = typeof global.showDirectoryPicker === 'function';
  var raiz = null;              // FileSystemDirectoryHandle
  var memoria = {};             // modo download: nome → texto

  function abrirBanco() {
    return new Promise(function (ok, erro) {
      var r = indexedDB.open(BANCO, 1);
      r.onupgradeneeded = function () { r.result.createObjectStore(LOJA); };
      r.onsuccess = function () { ok(r.result); };
      r.onerror = function () { erro(r.error); };
    });
  }

  function idb(modo, fn) {
    return abrirBanco().then(function (db) {
      return new Promise(function (ok, erro) {
        var tx = db.transaction(LOJA, modo);
        var req = fn(tx.objectStore(LOJA));
        tx.oncomplete = function () { db.close(); ok(req && req.result); };
        tx.onerror = function () { db.close(); erro(tx.error); };
      });
    });
  }

  function lembrar(handle) {
    return idb('readwrite', function (s) { return s.put(handle, 'pasta'); }).catch(function () {});
  }

  function esquecer() {
    return idb('readwrite', function (s) { return s.delete('pasta'); }).catch(function () {});
  }

  async function permissao(handle, pedir) {
    if (!handle) return 'denied';
    var opcoes = { mode: 'readwrite' };
    if (typeof handle.queryPermission !== 'function') return 'granted';
    var st = await handle.queryPermission(opcoes);
    if (st === 'granted' || !pedir) return st;
    return handle.requestPermission(opcoes);
  }

  /** Recupera a pasta usada na última vez. Retorna {nome, permissao} ou null. */
  async function restaurar() {
    if (!suportado) return null;
    var h;
    try { h = await idb('readonly', function (s) { return s.get('pasta'); }); } catch (e) { return null; }
    if (!h) return null;
    raiz = h;
    var st;
    try { st = await permissao(h, false); } catch (e) { st = 'prompt'; }
    return { nome: h.name, permissao: st };
  }

  /** Pede novamente a permissão de acesso à pasta lembrada (precisa de um clique do usuário). */
  async function reconectar() {
    var st = await permissao(raiz, true);
    if (st !== 'granted') throw new Error('Acesso à pasta não autorizado.');
    return raiz.name;
  }

  /** Abre o seletor de pastas do sistema. */
  async function escolher() {
    var h = await global.showDirectoryPicker({ id: 'processos-iptu', mode: 'readwrite' });
    var st = await permissao(h, true);
    if (st !== 'granted') throw new Error('Acesso à pasta não autorizado.');
    raiz = h;
    await lembrar(h);
    return h.name;
  }

  function conectada() { return !!raiz; }
  function nome() { return raiz ? raiz.name : ''; }

  async function listarJson() {
    var saida = [];
    if (suportado && raiz) {
      for await (var entrada of raiz.values()) {
        if (entrada.kind !== 'file' || !/\.json$/i.test(entrada.name)) continue;
        try {
          var f = await entrada.getFile();
          saida.push({ arquivo: entrada.name, texto: await f.text(), modificado: f.lastModified });
        } catch (e) {
          saida.push({ arquivo: entrada.name, erro: e.message });
        }
      }
    } else {
      Object.keys(memoria).forEach(function (n) {
        saida.push({ arquivo: n, texto: memoria[n].texto, modificado: memoria[n].modificado });
      });
    }
    return saida;
  }

  async function ler(arquivo) {
    if (suportado && raiz) {
      var fh = await raiz.getFileHandle(arquivo);
      return (await fh.getFile()).text();
    }
    if (!memoria[arquivo]) throw new Error('Arquivo não encontrado: ' + arquivo);
    return memoria[arquivo].texto;
  }

  async function existe(arquivo) {
    if (suportado && raiz) {
      try { await raiz.getFileHandle(arquivo); return true; } catch (e) { return false; }
    }
    return Object.prototype.hasOwnProperty.call(memoria, arquivo);
  }

  function baixar(arquivo, conteudo, tipo) {
    var blob = conteudo instanceof Blob ? conteudo : new Blob([conteudo], { type: tipo || 'application/json' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = arquivo;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  }

  async function gravarEm(dir, arquivo, conteudo) {
    var fh = await dir.getFileHandle(arquivo, { create: true });
    var w = await fh.createWritable();
    await w.write(conteudo);
    await w.close();
  }

  /** Grava um arquivo na pasta (ou baixa, no modo download). Retorna 'pasta' ou 'download'. */
  async function gravar(arquivo, texto) {
    if (suportado && raiz) {
      await gravarEm(raiz, arquivo, texto);
      return 'pasta';
    }
    memoria[arquivo] = { texto: texto, modificado: Date.now() };
    baixar(arquivo, texto, 'application/json');
    return 'download';
  }

  /** Grava a réplica (.docx) na subpasta "replicas" (ou baixa, no modo download). */
  async function gravarReplica(arquivo, blob) {
    if (suportado && raiz) {
      var sub = await raiz.getDirectoryHandle(SUBPASTA_REPLICAS, { create: true });
      await gravarEm(sub, arquivo, blob);
      return 'pasta';
    }
    baixar(arquivo, blob);
    return 'download';
  }

  async function remover(arquivo) {
    if (suportado && raiz) {
      await raiz.removeEntry(arquivo);
    } else {
      delete memoria[arquivo];
    }
  }

  /** Modo download: carrega os .json de uma lista de arquivos (input com webkitdirectory). */
  async function carregarArquivos(lista) {
    memoria = {};
    for (var i = 0; i < lista.length; i++) {
      var f = lista[i];
      var caminho = f.webkitRelativePath || f.name;
      // apenas os arquivos do primeiro nível da pasta
      if (caminho.split('/').length > 2 || !/\.json$/i.test(f.name)) continue;
      memoria[f.name] = { texto: await f.text(), modificado: f.lastModified };
    }
    var primeiro = lista[0] && lista[0].webkitRelativePath;
    return primeiro ? primeiro.split('/')[0] : '';
  }

  global.Pasta = {
    suportado: suportado,
    SUBPASTA_REPLICAS: SUBPASTA_REPLICAS,
    restaurar: restaurar,
    reconectar: reconectar,
    escolher: escolher,
    esquecer: function () { raiz = null; return esquecer(); },
    conectada: conectada,
    nome: nome,
    listarJson: listarJson,
    ler: ler,
    existe: existe,
    gravar: gravar,
    gravarReplica: gravarReplica,
    remover: remover,
    baixar: baixar,
    carregarArquivos: carregarArquivos
  };
})(window);
