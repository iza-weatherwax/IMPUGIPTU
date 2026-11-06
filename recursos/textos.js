/*
 * Montagem do texto da réplica fiscal a partir dos dados do processo.
 *
 * Convenções usadas nos textos:
 *   **texto**   → trecho em negrito
 *   ⟦rótulo⟧    → campo ainda não preenchido (destacado na prévia; "_____" no documento)
 *   {variavel}  → variável substituída pelo valor do processo (ver VARIAVEIS)
 */
(function (global) {
  'use strict';

  var NBSP = '\u00a0'; // espaço não separável: evita quebra de linha entre "nº" e o número, e entre o número e "m²"

  var MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho',
    'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];

  // Variáveis que o(a) fiscal pode usar nos textos livres.
  var VARIAVEIS = {
    protocolo: function (p) { return protocolo(p); },
    cci: function (p) { return limpo(p.processo.cci); },
    exercicio: function (p) { return limpo(p.processo.exercicio); },
    impugnante: function (p) { return limpo(p.requerente.nome); },
    area_lancada: function (p) { return area(p.analise.areaLancada); },
    area_apurada: function (p) { return area(p.analise.areaApurada); },
    area_confirmada: function (p) { return area(p.analise.areaConfirmada); }
  };

  function limpo(v) {
    return (v === undefined || v === null) ? '' : String(v).trim();
  }

  function pendente(valor, rotulo) {
    var v = limpo(valor);
    return v ? v : '⟦' + rotulo + '⟧';
  }

  function protocolo(p) {
    var n = limpo(p.processo.numero);
    var a = limpo(p.processo.ano);
    if (!n) return '';
    return a ? n + '/' + a : n;
  }

  // Converte "1.234,5" / "1234.5" / "1234,50 m²" em número; devolve NaN se não for possível.
  function numeroArea(v) {
    var s = limpo(v).replace(/m²|m2/gi, '').replace(/\s/g, '');
    if (!s) return NaN;
    if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
    else if ((s.match(/\./g) || []).length > 1) s = s.replace(/\./g, '');
    return /^\d+(\.\d+)?$/.test(s) ? Number(s) : NaN;
  }

  function area(v) {
    var n = numeroArea(v);
    if (isNaN(n)) return limpo(v);
    return n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + NBSP + 'm²';
  }

  function areaOuPendente(v, rotulo) {
    return limpo(v) ? area(v) : '⟦' + rotulo + '⟧';
  }

  function dataExtenso(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(limpo(iso));
    if (!m) return '';
    return Number(m[3]) + ' de ' + MESES[Number(m[2]) - 1] + ' de ' + m[1];
  }

  function dataCurta(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(limpo(iso));
    return m ? m[3] + '/' + m[2] + '/' + m[1] : '';
  }

  function resolverVariaveis(texto, p) {
    return String(texto || '').replace(/\{([a-z_]+)\}/g, function (todo, nome) {
      if (!VARIAVEIS[nome]) return todo;
      var v = VARIAVEIS[nome](p);
      return v ? v : '⟦' + nome + '⟧';
    });
  }

  function dividirParagrafos(texto) {
    return String(texto || '')
      .replace(/\r\n?/g, '\n')
      .split(/\n[ \t]*\n/)
      .map(function (s) { return s.replace(/^\s+|\s+$/g, ''); })
      .filter(Boolean);
  }

  // Ajustes tipográficos em todo o texto (inclusive o digitado pelo(a) fiscal):
  // mantém "nº" junto do número e o número junto de "m²".
  function tipografia(t) {
    return String(t)
      .replace(/(^|[\s(])([nN][º°]) (?=\S)/g, '$1$2' + NBSP)
      .replace(/(\d) (m²|m2)(?![\wÀ-ÿ])/g, '$1' + NBSP + '$2');
  }

  function semNegrito(t) { return String(t || '').replace(/\*\*/g, ''); }

  function comPonto(t) {
    t = limpo(t);
    return /[.!?:;]$/.test(t) ? t : t + '.';
  }

  function nomeDecisao(r) {
    return r === 'deferimento' ? 'DEFERIMENTO' : r === 'indeferimento' ? 'INDEFERIMENTO' : '';
  }

  function haDivergencia(p) {
    var a = p.analise.opiniaoTecnico, d = p.decisao.resultado;
    return !!(a && d && a !== d);
  }

  // Parágrafos gerados automaticamente a partir dos campos (texto do modelo).
  function paragrafosAutomaticos(p) {
    var cci = pendente(p.processo.cci, 'CCI');
    var pars = [];

    pars.push(comPonto('Uma vez recebido o processo de nº' + NBSP + pendente(protocolo(p), 'Nº DO PROCESSO') +
      ', restou identificado tratar-se de **IMPUGNAÇÃO** ao lançamento tributário do Imposto sobre a ' +
      'Propriedade Predial e Territorial Urbana (IPTU), referente ao ano-calendário de ' +
      pendente(p.processo.exercicio, 'EXERCÍCIO') + ' e ao imóvel de CCI nº' + NBSP + cci +
      ', pelo(a) impugnante ' + pendente(p.requerente.nome, 'NOME DO(A) IMPUGNANTE')));

    var a = p.analise;
    var opiniao = a.opiniaoTecnico === 'deferimento' ? 'deferimento'
      : a.opiniaoTecnico === 'indeferimento' ? 'indeferimento' : '⟦OPINIÃO DO(A) TÉCNICO(A)⟧';
    var constatacao;
    if (a.constatacao === 'confirmada') {
      constatacao = 'a **confirmação da área construída identificada, de ' +
        areaOuPendente(a.areaConfirmada, 'ÁREA CONFIRMADA') + ' para o CCI nº' + NBSP + comPonto(cci) + '**';
    } else if (a.constatacao === 'retificada') {
      constatacao = 'a **retificação da área construída, de ' +
        areaOuPendente(a.areaLancada, 'ÁREA LANÇADA') + ' para ' +
        areaOuPendente(a.areaApurada, 'ÁREA APURADA') + ', para o CCI nº' + NBSP + comPonto(cci) + '**';
    } else if (a.constatacao === 'outra') {
      var livre = semNegrito(resolverVariaveis(a.constatacaoTexto, p));
      constatacao = limpo(livre) ? '**' + comPonto(livre) + '**' : '⟦CONSTATAÇÃO DO RELATÓRIO⟧.';
    } else {
      constatacao = '⟦CONSTATAÇÃO DO RELATÓRIO TÉCNICO⟧.';
    }
    pars.push('Dando prosseguimento à análise, relata-se que esta auditoria fiscal teve acesso ao ' +
      'relatório emitido pelo(a) Técnico(a) Fiscal da Receita Municipal, o qual destaca sua opinião pelo ' +
      opiniao + ' da referida impugnação. Neste, consta ' + constatacao);

    dividirParagrafos(resolverVariaveis(a.paragrafoComplementar, p)).forEach(function (t) { pars.push(t); });

    if (haDivergencia(p)) {
      var just = dividirParagrafos(resolverVariaveis(p.decisao.justificativaDivergencia, p));
      if (just.length) just.forEach(function (t) { pars.push(t); });
      else pars.push('⟦JUSTIFICATIVA DA DIVERGÊNCIA EM RELAÇÃO AO RELATÓRIO TÉCNICO⟧');
    }

    var dec = nomeDecisao(p.decisao.resultado) || 'DECISÃO DO(A) FISCAL';
    pars.push('Dessa forma, no que diz respeito à competência deste Departamento de Fiscalização de ' +
      'Rendas/ISS, opina-se pelo ' + (p.decisao.resultado ? '**' + dec + '**' : '⟦' + dec + '⟧') +
      ' da referida impugnação.');

    pars.push('Sendo este o parecer técnico, encaminhando-se o presente processo à decisão da Gerência ' +
      'de Administração Tributária, Fiscalização e de Arrecadação, em consonância com o que enuncia o ' +
      'art. 581 da Lei Complementar 203/2022 (Código Tributário Municipal de Araguari/MG).');

    return pars;
  }

  function textoAutomatico(p) {
    return paragrafosAutomaticos(p).join('\n\n');
  }

  // Estrutura completa do documento, usada pela prévia, pela impressão e pelo .docx.
  function montarDocumento(p) {
    var corpo = p.replica && p.replica.personalizada
      ? dividirParagrafos(resolverVariaveis(p.replica.texto, p))
      : paragrafosAutomaticos(p);
    var f = p.fiscal || {};
    corpo = corpo.map(tipografia);
    return {
      numeroCabecalho: protocolo(p) || (limpo(p.processo.ano) ? ' /' + limpo(p.processo.ano) : ''),
      paragrafos: corpo,
      localData: p.decisao.incluirData
        ? 'Araguari/MG, ' + (dataExtenso(p.decisao.data) || '⟦DATA⟧') + '.'
        : '',
      espacoReduzido: !!(p.replica && p.replica.espacoReduzido),
      assinaturaNome: limpo(f.nome),
      assinaturaCargo: pendente(f.cargo, 'CARGO') + ' – Matrícula nº' + NBSP + pendente(f.matricula, 'MATRÍCULA')
    };
  }

  // Lista de campos obrigatórios ainda pendentes, por etapa do assistente.
  function pendencias(p) {
    var lista = [];
    function exige(cond, etapa, campo, texto) { if (!cond) lista.push({ etapa: etapa, campo: campo, texto: texto }); }
    exige(limpo(p.processo.numero), 0, 'processo.numero', 'Número do processo');
    exige(limpo(p.processo.ano), 0, 'processo.ano', 'Ano do processo');
    exige(limpo(p.processo.exercicio), 0, 'processo.exercicio', 'Exercício (ano-calendário) do IPTU');
    exige(limpo(p.processo.cci), 0, 'processo.cci', 'CCI do imóvel');
    exige(limpo(p.requerente.nome), 0, 'requerente.nome', 'Nome do(a) impugnante');

    var a = p.analise;
    exige(a.opiniaoTecnico, 1, 'analise.opiniaoTecnico', 'Opinião do(a) Técnico(a) Fiscal no relatório');
    exige(a.constatacao, 1, 'analise.constatacao', 'Constatação do relatório técnico');
    if (a.constatacao === 'confirmada') {
      exige(limpo(a.areaConfirmada), 1, 'analise.areaConfirmada', 'Área construída confirmada');
    } else if (a.constatacao === 'retificada') {
      exige(limpo(a.areaLancada), 1, 'analise.areaLancada', 'Área construída lançada');
      exige(limpo(a.areaApurada), 1, 'analise.areaApurada', 'Área construída apurada');
    } else if (a.constatacao === 'outra') {
      exige(limpo(a.constatacaoTexto), 1, 'analise.constatacaoTexto', 'Texto da constatação');
    }

    exige(p.decisao.resultado, 2, 'decisao.resultado', 'Decisão do(a) fiscal (deferimento ou indeferimento)');
    if (haDivergencia(p)) {
      exige(limpo(p.decisao.justificativaDivergencia), 2, 'decisao.justificativaDivergencia',
        'Justificativa da divergência em relação ao relatório técnico');
    }
    if (p.decisao.incluirData) exige(limpo(p.decisao.data), 2, 'decisao.data', 'Data da réplica');
    exige(limpo(p.fiscal.cargo), 2, 'fiscal.cargo', 'Cargo do(a) fiscal');
    exige(limpo(p.fiscal.matricula), 2, 'fiscal.matricula', 'Matrícula do(a) fiscal');

    if (p.replica && p.replica.personalizada) {
      exige(dividirParagrafos(p.replica.texto).length, 3, 'replica.texto', 'Texto personalizado da réplica');
      if (/⟦[^⟧]*⟧/.test(resolverVariaveis(p.replica.texto, p))) {
        lista.push({ etapa: 3, campo: 'replica.texto', texto: 'O texto personalizado contém campos pendentes' });
      }
    }
    return lista;
  }

  // Divide um parágrafo em trechos {texto, negrito}.
  function trechos(paragrafo) {
    var partes = String(paragrafo).split('**');
    if (partes.length > 1 && partes.length % 2 === 0) {
      // Quantidade ímpar de "**": o último é tratado como texto literal.
      var ult = partes.pop();
      partes[partes.length - 1] += '**' + ult;
    }
    return partes.map(function (t, i) { return { texto: t, negrito: i % 2 === 1 }; })
      .filter(function (t) { return t.texto !== ''; });
  }

  function paraTextoSimples(t) {
    return semNegrito(String(t)).replace(/⟦[^⟧]*⟧/g, '_____');
  }

  global.Textos = {
    VARIAVEIS: Object.keys(VARIAVEIS),
    montarDocumento: montarDocumento,
    textoAutomatico: textoAutomatico,
    pendencias: pendencias,
    haDivergencia: haDivergencia,
    resolverVariaveis: resolverVariaveis,
    trechos: trechos,
    paraTextoSimples: paraTextoSimples,
    protocolo: protocolo,
    area: area,
    dataCurta: dataCurta,
    dataExtenso: dataExtenso,
    nomeDecisao: nomeDecisao,
    limpo: limpo
  };
})(window);
