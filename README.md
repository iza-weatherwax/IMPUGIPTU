# Assistente de Réplica Fiscal — Impugnação de IPTU

Assistente local para preencher processos de impugnação de IPTU e emitir a **réplica fiscal** de
**deferimento** ou **indeferimento**, seguindo o modelo do Departamento de Fiscalização de Rendas/ISS
da Prefeitura Municipal de Araguari/MG.

- Tudo é feito no próprio computador, em uma página HTML: não há instalação, servidor, programa
  executável nem acesso à internet.
- Cada processo é salvo como um arquivo `.json` na pasta indicada pelo(a) fiscal.
- **Quem informa os dados e decide pelo deferimento ou indeferimento é o(a) fiscal responsável.** O
  assistente não sugere nem presume o resultado; ele só monta o texto conforme a decisão tomada.

## Como usar

1. Copie a pasta do assistente para o computador (ou para a rede) e abra o arquivo **`index.html`**
   no **Google Chrome** ou no **Microsoft Edge** (duplo clique ou arrastar para o navegador).
2. Clique em **Escolher pasta…** e indique a pasta onde os processos serão guardados. O navegador
   pede autorização para gravar nela. Nas próximas vezes, basta clicar em **Reconectar**.
3. Em **Dados do(a) fiscal**, confira cargo e matrícula (vêm preenchidos com os dados do modelo) e,
   se quiser, o nome. Esses dados entram na assinatura dos novos processos.
4. Clique em **+ Novo processo** e siga as 4 etapas:
   1. **Processo e impugnante** — número/ano do processo, exercício do IPTU, CCI e dados do(a) impugnante.
   2. **Relatório técnico** — opinião do(a) Técnico(a) Fiscal e o que consta no relatório
      (área confirmada, área retificada ou outra constatação). Parágrafos complementares são opcionais.
   3. **Decisão do(a) fiscal** — escolha **DEFERIMENTO** ou **INDEFERIMENTO**. Se a decisão divergir da
      opinião do relatório técnico, o assistente exige uma justificativa, que entra na réplica.
   4. **Revisão e emissão** — lista do que falta preencher, prévia do documento e emissão.
5. **Gerar réplica (.docx)** grava o documento Word na subpasta `replicas` e salva o processo.
   Também é possível **Imprimir / salvar em PDF** ou **Copiar texto**.

A prévia ao lado do formulário mostra a réplica em tempo real. Campos ainda não preenchidos aparecem
destacados em amarelo (no documento, saem como `_____`). Use **Ctrl+S** para salvar a qualquer momento.

## O que fica na pasta

```
Pasta escolhida/
├── IPTU_Processo_4521-2026.json           ← um arquivo por processo
├── IPTU_Processo_4600-2026.json
├── _configuracao-assistente.json          ← dados padrão do(a) fiscal
└── replicas/
    ├── Replica_Fiscal_4521-2026_INDEFERIMENTO.docx
    └── Replica_Fiscal_4600-2026_DEFERIMENTO.docx
```

- A lista inicial mostra todos os processos da pasta, com busca, filtros por decisão e situação e
  exportação para planilha (**Exportar lista (CSV)**).
- Processos sem número são salvos como `IPTU_Rascunho_xxxxxxxx.json` e renomeados quando o número é
  informado. Dois processos com o mesmo número/ano não são sobrescritos: o assistente avisa.
- Cada emissão fica registrada no próprio `.json` (data, decisão, arquivo gerado, assinatura e o texto
  emitido), formando o histórico do processo.

## Texto da réplica

O texto segue o modelo, mudando apenas o que depende dos dados e da decisão:

| Trecho | Indeferimento | Deferimento |
|---|---|---|
| Abertura | “Uma vez recebido o processo de nº …, restou identificado tratar-se de **IMPUGNAÇÃO** … ano-calendário de … CCI nº …, pelo(a) impugnante ….” | igual |
| Relatório técnico | “… destaca sua opinião pelo *(opinião do técnico)* … Neste, consta a **confirmação da área construída identificada, de … m² para o CCI nº ….**” | ou “… a **retificação da área construída, de … m² para … m², para o CCI nº ….**” |
| Conclusão | “… opina-se pelo **INDEFERIMENTO** da referida impugnação.” | “… opina-se pelo **DEFERIMENTO** …” |
| Encaminhamento | art. 581 da Lei Complementar 203/2022 | igual |

Opções adicionais, todas a critério do(a) fiscal:

- **Parágrafos complementares** (etapa 2), com sugestões prontas que podem ser editadas.
- **Justificativa da divergência**, obrigatória quando a decisão difere da opinião técnica.
- **Local e data** (“Araguari/MG, 6 de novembro de 2026.”) — o modelo original não traz data.
- **Nome do(a) fiscal** acima do cargo — o modelo original traz apenas cargo e matrícula.
- **Reduzir o espaço antes da assinatura**, para textos longos que não caibam em uma página.
- **Editar o texto manualmente** (etapa 4), para ajustes pontuais de redação.

Nos textos livres:

- `**texto**` fica em **negrito**;
- linha em branco separa parágrafos;
- as variáveis `{cci}`, `{protocolo}`, `{exercicio}`, `{impugnante}`, `{area_lancada}`, `{area_apurada}` e
  `{area_confirmada}` são trocadas pelos valores do processo.

## Outros navegadores

No Firefox e no Safari, que não permitem gravar diretamente em uma pasta, o assistente funciona em
**modo de compatibilidade**: use **Carregar processos de uma pasta…** para ver os processos já
existentes; ao salvar ou emitir, os arquivos são baixados e devem ser guardados na pasta dos processos.

Evite abrir o assistente em janela anônima/InPrivate: nela, o navegador não lembra a pasta escolhida.

## Arquivos do assistente

| Arquivo | Conteúdo |
|---|---|
| `index.html` | Página do assistente (abra este arquivo). |
| `recursos/app.js` | Interface: lista de processos, etapas, prévia, salvamento e emissão. |
| `recursos/textos.js` | Montagem do texto da réplica (deferimento/indeferimento) e verificação de pendências. |
| `recursos/docx.js` | Geração do arquivo Word (.docx) no próprio navegador. |
| `recursos/pasta.js` | Leitura e gravação dos arquivos na pasta escolhida. |
| `recursos/modelo-replica.js` | Partes do modelo .docx original (cabeçalho, brasão, marca-d’água e estilos). |
| `recursos/estilo.css` | Aparência da página e da folha A4 da prévia/impressão. |

## Formato do arquivo `.json`

```json
{
  "formato": "impugnacao-iptu/processo",
  "versao": 1,
  "id": "…",
  "criadoEm": "2026-11-06T12:00:00.000Z",
  "atualizadoEm": "2026-11-06T12:30:00.000Z",
  "situacao": "rascunho | emitida",
  "processo": { "numero": "4521", "ano": "2026", "protocolo": "4521/2026", "dataEntrada": "2026-03-14",
                "exercicio": "2026", "cci": "01.02.0345.0012", "enderecoImovel": "" },
  "requerente": { "nome": "", "documento": "", "contato": "", "alegacoes": "" },
  "analise": { "tecnicoNome": "", "dataRelatorio": "", "opiniaoTecnico": "deferimento | indeferimento",
               "constatacao": "confirmada | retificada | outra", "areaConfirmada": "", "areaLancada": "",
               "areaApurada": "", "constatacaoTexto": "", "areaAlegada": "",
               "paragrafoComplementar": "", "observacoes": "" },
  "decisao": { "resultado": "deferimento | indeferimento", "justificativaDivergencia": "",
               "data": "2026-11-06", "incluirData": false },
  "fiscal": { "nome": "", "cargo": "Auditora Fiscal da Receita Municipal", "matricula": "91979" },
  "replica": { "personalizada": false, "texto": "", "espacoReduzido": false,
               "emissoes": [ { "em": "…", "resultado": "indeferimento", "arquivo": "Replica_Fiscal_….docx",
                               "fiscal": { }, "texto": [ "…parágrafos emitidos…" ] } ] }
}
```

Campos marcados como “registro interno” (alegações, área alegada, observações, dados do técnico) ficam
apenas no `.json` e não aparecem na réplica.
