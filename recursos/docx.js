/*
 * Geração do arquivo .docx da réplica fiscal, 100% no navegador.
 * Usa as partes do modelo (recursos/modelo-replica.js) e monta apenas o corpo do documento.
 * O pacote .docx é um ZIP; aqui ele é gravado sem compressão (método "store"), aceito pelo Word.
 */
(function (global) {
  'use strict';

  // ---------- ZIP ----------
  var TABELA_CRC = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = TABELA_CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function dataDos(d) {
    return {
      hora: ((d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1)) & 0xFFFF,
      data: (((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xFFFF
    };
  }

  function zip(arquivos) {
    var enc = new TextEncoder();
    var agora = dataDos(new Date());
    var locais = [], centrais = [], deslocamento = 0;

    arquivos.forEach(function (a) {
      var nome = enc.encode(a.nome);
      var dados = a.dados;
      var crc = crc32(dados);

      var lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true);
      lh.setUint16(4, 20, true);
      lh.setUint16(6, 0x0800, true);      // nomes em UTF-8
      lh.setUint16(8, 0, true);           // sem compressão
      lh.setUint16(10, agora.hora, true);
      lh.setUint16(12, agora.data, true);
      lh.setUint32(14, crc, true);
      lh.setUint32(18, dados.length, true);
      lh.setUint32(22, dados.length, true);
      lh.setUint16(26, nome.length, true);
      lh.setUint16(28, 0, true);
      locais.push(new Uint8Array(lh.buffer), nome, dados);

      var ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true);
      ch.setUint16(4, 20, true);
      ch.setUint16(6, 20, true);
      ch.setUint16(8, 0x0800, true);
      ch.setUint16(10, 0, true);
      ch.setUint16(12, agora.hora, true);
      ch.setUint16(14, agora.data, true);
      ch.setUint32(16, crc, true);
      ch.setUint32(20, dados.length, true);
      ch.setUint32(24, dados.length, true);
      ch.setUint16(28, nome.length, true);
      ch.setUint32(42, deslocamento, true);
      centrais.push(new Uint8Array(ch.buffer), nome);

      deslocamento += 30 + nome.length + dados.length;
    });

    var tamCentral = centrais.reduce(function (s, b) { return s + b.length; }, 0);
    var fim = new DataView(new ArrayBuffer(22));
    fim.setUint32(0, 0x06054b50, true);
    fim.setUint16(8, arquivos.length, true);
    fim.setUint16(10, arquivos.length, true);
    fim.setUint32(12, tamCentral, true);
    fim.setUint32(16, deslocamento, true);

    return new Blob(locais.concat(centrais, [new Uint8Array(fim.buffer)]), {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
  }

  function base64ParaBytes(b64) {
    var bin = atob(b64);
    var out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  // ---------- WordprocessingML ----------
  function esc(t) {
    return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function rPr(tamanho, negrito) {
    return '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>' +
      (negrito ? '<w:b/><w:bCs/>' : '') +
      '<w:sz w:val="' + tamanho + '"/><w:szCs w:val="' + tamanho + '"/><w:lang w:eastAsia="pt-BR"/></w:rPr>';
  }

  // Converte o texto (com **negrito** e quebras de linha) em runs do Word.
  function runs(texto, tamanho, quebraInicial) {
    var Textos = global.Textos;
    var xml = '';
    var primeiro = true;
    Textos.trechos(texto).forEach(function (tr) {
      var linhas = Textos.paraTextoSimples(tr.texto).split('\n');
      var conteudo = '';
      if (primeiro && quebraInicial) conteudo += '<w:br/>';
      linhas.forEach(function (l, i) {
        if (i > 0) conteudo += '<w:br/>';
        if (l !== '') conteudo += '<w:t xml:space="preserve">' + esc(l) + '</w:t>';
      });
      xml += '<w:r>' + rPr(27, tr.negrito) + conteudo + '</w:r>';
      primeiro = false;
    });
    if (!xml && quebraInicial) xml = '<w:r>' + rPr(27, false) + '<w:br/></w:r>';
    return xml;
  }

  var PPR_CORPO = '<w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="705"/><w:jc w:val="both"/></w:pPr>';
  var PPR_DIREITA = '<w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/><w:jc w:val="right"/></w:pPr>';

  function paragrafoVazio(linha, depois) {
    return '<w:p><w:pPr><w:spacing w:after="' + depois + '" w:line="' + linha + '" w:lineRule="auto"/>' +
      (linha === 360 ? '<w:ind w:firstLine="705"/><w:jc w:val="both"/>' : '') + '</w:pPr>' +
      '<w:r>' + rPr(24, false) + '<w:t xml:space="preserve"> </w:t></w:r></w:p>';
  }

  function linhaAssinatura(texto, depois) {
    return '<w:p><w:pPr><w:spacing w:after="' + depois + '" w:line="256" w:lineRule="auto"/><w:jc w:val="right"/></w:pPr>' +
      '<w:r>' + rPr(24, false) + '<w:t xml:space="preserve">' + esc(global.Textos.paraTextoSimples(texto)) + '</w:t></w:r></w:p>';
  }

  function corpoDocumento(doc) {
    var xml = '';
    doc.paragrafos.forEach(function (p, i) {
      // Como no modelo: a partir do 2º parágrafo, uma quebra de linha inicial cria o espaço entre parágrafos.
      xml += '<w:p>' + PPR_CORPO + runs(p, 27, i > 0) + '</w:p>';
    });
    if (doc.localData) xml += '<w:p>' + PPR_DIREITA + runs(doc.localData, 27, true) + '</w:p>';

    // Espaço para assinatura: o mesmo do modelo (cerca de 6 cm) ou reduzido, a critério do(a) fiscal.
    var i;
    if (doc.espacoReduzido) {
      xml += paragrafoVazio(360, 0);
    } else {
      for (i = 0; i < 3; i++) xml += paragrafoVazio(360, 0);
      for (i = 0; i < 6; i++) xml += paragrafoVazio(240, 0);
    }
    xml += paragrafoVazio(256, 165);

    if (doc.assinaturaNome) xml += linhaAssinatura(doc.assinaturaNome, 0);
    xml += linhaAssinatura(doc.assinaturaCargo, 165);
    return xml;
  }

  function propriedades(titulo, autor) {
    var agora = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + esc(titulo) + '</dc:title><dc:subject>Impugnação de IPTU</dc:subject>' +
      '<dc:creator>' + esc(autor) + '</dc:creator><cp:lastModifiedBy>' + esc(autor) + '</cp:lastModifiedBy>' +
      '<cp:revision>1</cp:revision>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + agora + '</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">' + agora + '</dcterms:modified></cp:coreProperties>';
  }

  /**
   * Gera o .docx (Blob) a partir da estrutura devolvida por Textos.montarDocumento().
   */
  function gerar(doc, titulo, autor) {
    var M = global.MODELO_REPLICA;
    if (!M) throw new Error('Modelo da réplica não carregado (recursos/modelo-replica.js).');
    var enc = new TextEncoder();
    var arquivos = [];
    function txt(nome, conteudo) { arquivos.push({ nome: nome, dados: enc.encode(conteudo) }); }

    var ordem = ['[Content_Types].xml', '_rels/.rels', 'docProps/app.xml'];
    ordem.forEach(function (n) { txt(n, M.partes[n]); });
    txt('docProps/core.xml', propriedades(titulo, autor));
    txt('word/document.xml', M.documentoInicio + corpoDocumento(doc) + M.documentoFim);

    Object.keys(M.partes).forEach(function (n) {
      if (ordem.indexOf(n) >= 0) return;
      var conteudo = M.partes[n];
      if (n === 'word/header2.xml') {
        conteudo = conteudo.replace('{{NUMERO_PROCESSO}}', esc(doc.numeroCabecalho || ''));
      }
      txt(n, conteudo);
    });

    var imagens = {};
    Object.keys(M.imagens).forEach(function (n) {
      imagens[n] = base64ParaBytes(M.imagens[n]);
      arquivos.push({ nome: n, dados: imagens[n] });
    });
    Object.keys(M.copias).forEach(function (n) {
      arquivos.push({ nome: n, dados: imagens[M.copias[n]] });
    });
    return zip(arquivos);
  }

  global.Docx = { gerar: gerar, zip: zip, crc32: crc32 };
})(window);
