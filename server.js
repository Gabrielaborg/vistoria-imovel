#!/usr/bin/env node
const http = require('http');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const { execSync } = require('child_process');

const PORT = 3000;
const OUTPUT_DIR = path.join(__dirname, 'laudos_gerados');
const COVER_IMG = path.join(__dirname, 'cover_page.png');
const FOOTER_IMG = path.join(__dirname, 'footer_bar.jpeg');
const LOGO_IMG = path.join(__dirname, 'logo.png');
const HISTORICO_FILE = path.join(OUTPUT_DIR, 'historico.json');
const AGENDA_FILE = path.join(OUTPUT_DIR, 'agendamentos.json');

// Chave de API da Anthropic, lida de variável de ambiente (NUNCA fica no código nem no navegador).
// Configure em Railway: Service > Variables > ANTHROPIC_API_KEY
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || '';

// Chave de API do Geoapify (gratuita), usada para buscar a imagem do mapa pela localização atual.
// Configure em Railway: Service > Variables > GEOAPIFY_API_KEY
// Como conseguir de graça: https://www.geoapify.com/ → Sign Up → cria um projeto → copia a API Key.
const GEOAPIFY_API_KEY = process.env.GEOAPIFY_API_KEY || '';

// Senha de acesso ao app inteiro (protege dados de clientes: CPF, endereço, etc).
// Configure em Railway: Service > Variables > APP_PASSWORD
// Se não configurada, o app fica sem senha (comportamento de antes, sem travar nada).
const APP_PASSWORD = process.env.APP_PASSWORD || '';

const SEIS_MESES_MS = 6 * 30 * 24 * 60 * 60 * 1000; // aproximação de 6 meses, usada pra decidir o que é "antigo"

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

function lerHistorico() { try { return JSON.parse(fs.readFileSync(HISTORICO_FILE, 'utf8')); } catch { return []; } }
function salvarHistorico(h) { fs.writeFileSync(HISTORICO_FILE, JSON.stringify(h, null, 2)); }
function lerAgenda() { try { return JSON.parse(fs.readFileSync(AGENDA_FILE, 'utf8')); } catch { return []; } }
function salvarAgenda(a) { fs.writeFileSync(AGENDA_FILE, JSON.stringify(a, null, 2)); }

// ── Compactação de laudos antigos (economiza espaço no Volume do Railway) ──

// Compacta um arquivo com gzip (reduz bastante o tamanho, principalmente do .docx,
// que já é um zip por dentro mas ainda tem folga) e apaga o original.
function compactarArquivo(caminhoOriginal) {
  const dados = fs.readFileSync(caminhoOriginal);
  const comprimido = zlib.gzipSync(dados, { level: 9 });
  fs.writeFileSync(caminhoOriginal + '.gz', comprimido);
  fs.unlinkSync(caminhoOriginal);
}

// Descompacta de volta pra memória na hora do download, sem precisar guardar
// uma cópia solta do arquivo original no disco.
function lerArquivoTalvezCompactado(caminhoOriginal) {
  if (fs.existsSync(caminhoOriginal)) return fs.readFileSync(caminhoOriginal);
  if (fs.existsSync(caminhoOriginal + '.gz')) return zlib.gunzipSync(fs.readFileSync(caminhoOriginal + '.gz'));
  return null;
}
function arquivoExisteOuCompactado(caminhoOriginal) {
  return fs.existsSync(caminhoOriginal) || fs.existsSync(caminhoOriginal + '.gz');
}

// Roda periodicamente: procura laudos com mais de 6 meses no histórico e compacta
// o Word/PDF deles (ficam bem menores, mas continuam disponíveis pra download normal
// no app — só demoram uma fração de segundo a mais, pela descompactação na hora).
function compactarLaudosAntigos() {
  const hist = lerHistorico();
  let mudou = false;
  for (const item of hist) {
    if (item.compactado) continue;
    if (!item.ts || (Date.now() - item.ts) < SEIS_MESES_MS) continue;
    try {
      const docxPath = path.join(OUTPUT_DIR, item.arquivo);
      const pdfPath = docxPath.replace('.docx', '.pdf');
      let compactouAlgo = false;
      if (fs.existsSync(docxPath)) { compactarArquivo(docxPath); compactouAlgo = true; }
      if (fs.existsSync(pdfPath)) { compactarArquivo(pdfPath); compactouAlgo = true; }
      if (compactouAlgo) {
        item.compactado = true;
        mudou = true;
        console.log(`✓ Laudo compactado (economia de espaço): ${item.arquivo}`);
      }
    } catch (e) {
      console.error(`Falha ao compactar "${item.arquivo}", tenta de novo na próxima rodada:`, e.message);
    }
  }
  if (mudou) salvarHistorico(hist);
}

setTimeout(() => { try { compactarLaudosAntigos(); } catch (e) { console.error('Erro na rotina de compactação:', e.message); } }, 30 * 1000);
setInterval(() => { try { compactarLaudosAntigos(); } catch (e) { console.error('Erro na rotina de compactação:', e.message); } }, 24 * 60 * 60 * 1000);

// ── Chamada central à API da Anthropic (roda só no servidor, chave nunca exposta) ──
async function chamarClaude({ system, messages, maxTokens = 300 }) {
  if (!ANTHROPIC_API_KEY) throw new Error('ANTHROPIC_API_KEY não configurada no servidor.');
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-5',
      max_tokens: maxTokens,
      system,
      messages
    })
  });
  if (!res.ok) {
    const errText = await res.text().catch(() => '');
    throw new Error(`Anthropic API respondeu ${res.status}: ${errText}`);
  }
  const data = await res.json();
  return data.content?.[0]?.text || '';
}

// Lê o corpo (body) de uma requisição POST como JSON
function lerCorpoJSON(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try { resolve(JSON.parse(body)); } catch (e) { reject(e); }
    });
    req.on('error', reject);
  });
}

async function gerarDocx(payload) {
  const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, Footer } = require('docx');
  const { dados, tipoVistoria, obsGeral, registros, plantaBase64, plantaMediaType, mapaBase64, mapaMediaType, conclusaoIA } = payload;

  const dataFmt = dados.data
    ? new Date(dados.data + 'T12:00:00').toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' })
    : new Date().toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' });

  const coverImgData = fs.readFileSync(COVER_IMG);
  const footerImgData = fs.readFileSync(FOOTER_IMG);
  const logoImgData = fs.readFileSync(LOGO_IMG);

  const FONT = 'Montserrat';

  // ─── Helpers ────────────────────────────────────────────
  const B = (text, size=20) => new TextRun({ text, bold: true, size, font: FONT });
  const N = (text, size=20) => new TextRun({ text, size, font: FONT });
  const br = () => new Paragraph({ children: [N('')], spacing: { after: 120 } });
  const indent = { left: 720 };

  const secNum = (n, text) => new Paragraph({
    children: [B(`${n}. ${text}`, 22)],
    spacing: { before: 300, after: 140 }
  });
  const subNum = (n, text) => new Paragraph({
    children: [B(`${n} `, 20), B(text, 20)],
    indent, spacing: { before: 180, after: 80 }
  });
  const bodyP = (text) => new Paragraph({
    children: [N(text, 20)], indent, spacing: { after: 100 }, alignment: AlignmentType.JUSTIFIED
  });
  const bodyBold = (label, val) => new Paragraph({
    children: [B(label, 20), N(val, 20)], indent, spacing: { after: 80 }
  });
  const bulletP = (text) => new Paragraph({
    children: [N(`• ${text}`, 20)], indent: { left: 1260 }, spacing: { after: 60 }
  });
  const defP = (label, text) => new Paragraph({
    children: [B(label+': ', 20), N(text, 20)], indent, spacing: { after: 80 }, alignment: AlignmentType.JUSTIFIED
  });

  // ─── Footer ─────────────────────────────────────────────
  const makeFooter = () => new Footer({
    children: [
      new Paragraph({
        children: [new ImageRun({ data: footerImgData, transformation: { width: 520, height: 17 }, type: 'jpg' })],
        alignment: AlignmentType.CENTER, spacing: { before: 60 }
      })
    ]
  });

  // ─── Registros ──────────────────────────────────────────
  const registrosParagraphs = [];
  let imgCounter = 1; // numeração contínua por TODO o laudo (não reinicia a cada ambiente/defeito)
  const byAmbiente = {};
  registros.forEach(r => { if (!byAmbiente[r.ambiente]) byAmbiente[r.ambiente] = []; byAmbiente[r.ambiente].push(r); });

  for (const [ambiente, items] of Object.entries(byAmbiente)) {
    registrosParagraphs.push(new Paragraph({
      children: [B(ambiente, 20)], indent, spacing: { before: 200, after: 100 }
    }));
    // Agrupa registros com o MESMO texto de defeito dentro do mesmo ambiente,
    // mesmo que tenham sido salvos em momentos separados durante a vistoria.
    const porDefeito = {};
    const ordemDefeitos = [];
    for (const item of items) {
      if (!porDefeito[item.defeito]) { porDefeito[item.defeito] = []; ordemDefeitos.push(item.defeito); }
      porDefeito[item.defeito].push(...item.fotos);
    }
    for (const defeito of ordemDefeitos) {
      const fotosDoGrupo = porDefeito[defeito];
      // Texto completo do defeito aparece UMA VEZ por grupo (não repete por foto nem por registro separado)
      registrosParagraphs.push(new Paragraph({
        children: [N(defeito, 20)], indent, spacing: { after: 80 }, alignment: AlignmentType.JUSTIFIED
      }));
      const totalFotos = fotosDoGrupo.length;
      const inicio = imgCounter;
      const fim = imgCounter + totalFotos - 1;
      const referencia = totalFotos > 1 ? `Evidenciado nas imagens ${inicio} a ${fim}.` : `Evidenciado na imagem ${inicio}.`;
      registrosParagraphs.push(new Paragraph({
        children: [N(referencia, 18)], indent, spacing: { after: 100 }
      }));
      for (let i = 0; i < totalFotos; i++) {
        const foto = fotosDoGrupo[i];
        registrosParagraphs.push(new Paragraph({
          children: [N(`Imagem ${imgCounter}`, 20)], alignment: AlignmentType.CENTER, spacing: { after: 60 }
        }));
        try {
          const imgBuf = Buffer.from(foto.base64, 'base64');
          const imgType = foto.mediaType === 'image/png' ? 'png' : 'jpg';
          registrosParagraphs.push(new Paragraph({
            children: [new ImageRun({ data: imgBuf, transformation: { width: 400, height: 280 }, type: imgType })],
            alignment: AlignmentType.CENTER, spacing: { before: 80, after: 160 }
          }));
        } catch(e) { console.error('Img error:', e.message); }
        imgCounter++;
      }
    }
    registrosParagraphs.push(br());
  }

  // ─── Planta ─────────────────────────────────────────────
  const plantaParagraphs = [];
  if (plantaBase64) {
    try {
      const plantaBuf = Buffer.from(plantaBase64, 'base64');
      const plantaType = plantaMediaType === 'image/png' ? 'png' : 'jpg';
      plantaParagraphs.push(
        secNum(5, 'PLANTA DO IMÓVEL'),
        bodyP('Abaixo, apresenta-se planta semelhante da unidade vistoriada:'),
        new Paragraph({
          children: [new ImageRun({ data: plantaBuf, transformation: { width: 420, height: 300 }, type: plantaType })],
          alignment: AlignmentType.CENTER, spacing: { before: 100, after: 160 }
        }),
        br()
      );
    } catch(e) { console.error('Planta error:', e.message); }
  }

  // ─── Mapa ────────────────────────────────────────────────
  const mapaParagraphs = [];
  if (mapaBase64) {
    try {
      const mapaBuf = Buffer.from(mapaBase64, 'base64');
      const mapaType = mapaMediaType === 'image/png' ? 'png' : 'jpg';
      mapaParagraphs.push(
        new Paragraph({ children: [B('Localização:', 20)], indent, spacing: { before: 100, after: 60 } }),
        new Paragraph({
          children: [new ImageRun({ data: mapaBuf, transformation: { width: 420, height: 220 }, type: mapaType })],
          alignment: AlignmentType.CENTER, spacing: { before: 60, after: 120 }
        })
      );
    } catch(e) { console.error('Mapa error:', e.message); }
  } else {
    mapaParagraphs.push(new Paragraph({ children: [B('Localização:', 20), N(' ' + [dados.endereco, dados.bloco, dados.apto, dados.cidade].filter(Boolean).join(', '), 20)], indent, spacing: { after: 80 } }));
  }

  // ─── Cômodos ────────────────────────────────────────────
  const comodosLines = (dados.comodos||'').split('\n').map(l=>l.trim()).filter(Boolean);
  const comodosItems = comodosLines.length > 0 ? comodosLines.map(l => bulletP(l.replace(/[;,.]$/,'') + ';')) : [bulletP('Conforme especificações do projeto.')];

  const endFull = [dados.endereco, dados.bloco, dados.apto, dados.cidade, dados.cep].filter(Boolean).join(', ');
  const numOffset = plantaParagraphs.length > 0 ? 1 : 0;

  // ─── Textos fixos ────────────────────────────────────────
  const elaboracaoTexto1 = 'A elaboração do presente relatório de vistoria técnica de recebimento da unidade habitacional foi realizada com base na identificação dos elementos construtivos aparentes, sua localização dentro do imóvel e as manifestações patológicas visíveis no momento da inspeção.';
  const elaboracaoTexto2 = 'Durante a vistoria, foram observados diversos pontos de não conformidade, falhas de acabamento, anomalias e possíveis vícios construtivos que podem comprometer o desempenho esperado dos sistemas e materiais.';
  const elaboracaoTexto3 = 'A inspeção foi feita com base nos princípios estabelecidos pela ABNT NBR 16747:2020 – Diretrizes para inspeção predial, na ABNT NBR 5674:2024 – Manutenção de edificações, e também conforme os conceitos definidos pelo IBAPE Nacional.';
  const elaboracaoTexto4 = 'Considerando que alguns termos utilizados neste documento podem não ser de conhecimento geral, seguem abaixo os principais conceitos utilizados ao longo do relatório:';
  const elaboracaoTexto5 = 'A unidade inspecionada apresenta diversas não conformidades visuais. Tais ocorrências indicam ausência de cuidados na execução final e comprometem o recebimento do imóvel em condições ideais de entrega.';
  const elaboracaoTexto6 = 'A recomendação técnica é que todas as anomalias listadas neste relatório sejam corrigidas antes da conclusão da entrega da unidade ao proprietário, garantindo o desempenho mínimo esperado e evitando prejuízos futuros.';

  // Se o front-end mandou uma conclusão gerada por IA (baseada nos defeitos reais), usa ela.
  // Caso contrário (IA falhou, ou não foi chamada), cai no texto fixo genérico de sempre.
  const conclusaoTextoFixo = [
    'Com base na vistoria técnica realizada na unidade habitacional, constatou-se a presença de não conformidades construtivas, falhas de acabamento e inconformidades funcionais distribuídas nos ambientes inspecionados, conforme descrito e documentado ao longo deste relatório técnico. As manifestações observadas incluem irregularidades em revestimentos, falhas de rejuntamento, defeitos em pintura, problemas em esquadrias, portas, elementos hidráulicos, acabamentos e demais sistemas construtivos aparentes.',
    'Os defeitos identificados evidenciam deficiência nos processos executivos e no controle de qualidade durante as etapas de acabamento e entrega da unidade, não sendo compatíveis com o padrão esperado para um imóvel novo. Ainda que parte das inconformidades apresente caráter predominantemente estético, diversas manifestações podem comprometer a durabilidade dos materiais, o desempenho dos sistemas construtivos, a estanqueidade, a funcionalidade dos ambientes e a vida útil da edificação ao longo do tempo.',
    'Conforme os princípios estabelecidos pela ABNT NBR 15575, a edificação deve atender aos requisitos mínimos de desempenho relacionados à segurança, habitabilidade, funcionalidade e durabilidade. Da mesma forma, os serviços executivos e acabamentos devem seguir padrões adequados de qualidade e conformidade técnica, observando as boas práticas construtivas e as normas aplicáveis a cada sistema construtivo. As anomalias constatadas neste relatório demonstram inconformidades em relação a tais requisitos, tornando tecnicamente recomendável a correção integral dos itens apontados.',
    'Dessa forma, conclui-se que todas as não conformidades registradas neste documento devem ser devidamente corrigidas pela construtora/responsável técnico antes da aceitação definitiva do imóvel, garantindo o adequado desempenho dos sistemas, a preservação da vida útil dos materiais e o padrão de qualidade esperado para a edificação. Recomenda-se ainda que os reparos sejam executados com acompanhamento técnico e observância aos procedimentos normativos aplicáveis, a fim de evitar recorrência das falhas identificadas.'
  ];
  const conclusaoTexto = (conclusaoIA && conclusaoIA.trim())
    ? conclusaoIA.trim().split(/\n\s*\n/).map(p => p.trim()).filter(Boolean)
    : conclusaoTextoFixo;

  const doc = new Document({
    sections: [
      // CAPA
      {
        properties: { page: { size: { width: 11910, height: 16840 }, margin: { top: 0, right: 0, bottom: 0, left: 0 } } },
        children: [
          new Paragraph({
            children: [new ImageRun({ data: coverImgData, transformation: { width: 794, height: 1123 }, type: 'png' })],
            alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }
          })
        ]
      },
      // CONTEÚDO
      {
        properties: { page: { size: { width: 11910, height: 16840 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
        footers: { default: makeFooter() },
        children: [
          // 1. IDENTIFICAÇÃO
          secNum(1, 'IDENTIFICAÇÃO DO CONTRATANTE'),
          bodyP(`${dados.nome || 'Cliente'}, portador(a) do CPF nº ${dados.cpf || '—'}${dados.telefone ? `, telefone ${dados.telefone}` : ''}, solicitou a elaboração do presente relatório de vistoria de imóvel, com o objetivo de registrar as condições da unidade no momento da entrega, identificando eventuais inconformidades aparentes e falhas de execução visíveis.`),
          br(),

          // 2. OBJETIVO
          secNum(2, 'OBJETIVO'),
          bodyP(`Este relatório tem como finalidade documentar, de forma objetiva e detalhada, as condições do imóvel ${tipoVistoria === 'Imóvel Novo' ? 'novo ' : ''}na data da vistoria, identificando eventuais não conformidades em acabamentos, instalações elétricas e hidráulicas, estrutura e funcionalidade dos ambientes. A avaliação foi conduzida seguindo as diretrizes estabelecidas pelas normas técnicas vigentes, incluindo as NBR (Normas Brasileiras) e os referenciais do PBQP-H (Programa Brasileiro da Qualidade e Produtividade no Habitat), assegurando que os padrões de qualidade, segurança, funcionalidade e durabilidade do empreendimento sejam observados.`),
          br(),

          // 3. DADOS INICIAIS
          secNum(3, 'DADOS INICIAIS'),
          subNum('3.1', 'Identificação'),
          bodyBold('Empreendimento: ', dados.empreendimento || '—'),
          bodyBold('Endereço: ', endFull),
          ...mapaParagraphs,
          br(),
          subNum('3.2', 'Realização da vistoria'),
          bodyP('Responsável: Engenheira Civil Gabriela Soares Borges'),
          bodyP('Registro CREA: 427760MG'),
          br(),
          subNum('3.3', 'Data das Vistorias'),
          bodyP(`A vistoria foi realizada dia ${dataFmt}.`),
          br(),

          // 4. DESCRIÇÃO DO IMÓVEL
          secNum(4, 'DESCRIÇÃO DO IMÓVEL'),
          bodyP('O imóvel vistoriado trata-se de um apartamento com a seguinte configuração:'),
          ...(dados.metragem ? [bulletP(`Área total: ${dados.metragem} m²`)] : []),
          ...comodosItems,
          br(),
          bodyP('Durante a vistoria, foram inspecionados os acabamentos, instalações elétricas e hidráulicas, funcionalidade dos ambientes e demais itens que compõem o imóvel, registrando-se eventuais não conformidades para que sejam corrigidas conforme os padrões de qualidade estabelecidos pela construtora.'),
          br(),

          // 5. PLANTA (se houver)
          ...plantaParagraphs,

          // 6. ELABORAÇÃO DE RELATÓRIO
          secNum(5 + numOffset, 'ELABORAÇÃO DE RELATÓRIO'),
          bodyP(elaboracaoTexto1),
          br(),
          bodyP(elaboracaoTexto2),
          br(),
          bodyP(elaboracaoTexto3),
          br(),
          bodyP(elaboracaoTexto4),
          br(),
          defP('Anomalia', 'Irregularidade que compromete o desempenho de um elemento ou sistema da edificação. Pode ter origem no projeto, execução, uso ou manutenção inadequada.'),
          defP('Manifestação Patológica', 'Sinais visíveis de degradação, como fissuras, manchas, destacamentos, entre outros.'),
          defP('Agente de Degradação', 'Fatores (naturais, físicos ou químicos) que contribuem para a deterioração dos elementos construtivos.'),
          defP('Falha', 'Perda da função de um componente, seja por uso indevido, má execução ou falta de manutenção.'),
          defP('Desempenho', 'Comportamento da edificação e seus sistemas durante o uso, frente às solicitações normais esperadas ao longo de sua vida útil.'),
          defP('Vida Útil (VU)', 'Período em que um sistema ou componente deve cumprir suas funções, conforme previsto em projeto e respeitada sua manutenção adequada.'),
          defP('Plano de Manutenção', 'Documento técnico que organiza as ações necessárias de manutenção preventiva e corretiva de uma edificação.'),
          br(),
          bodyP(elaboracaoTexto5),
          br(),
          bodyP(elaboracaoTexto6),
          br(),

          // 7. REGISTROS
          secNum(6 + numOffset, 'REGISTRO DE NÃO CONFORMIDADES DA VISTORIA'),
          bodyP(`A seguir, são apresentados os registros fotográficos das não conformidades identificadas durante a vistoria no dia ${dataFmt}, acompanhados da respectiva descrição detalhada.`),
          br(),
          ...registrosParagraphs,

          // OUTROS PROBLEMAS
          ...(obsGeral ? [
            secNum(7 + numOffset, 'OUTROS PROBLEMAS'),
            ...obsGeral.split('\n').map(p => bodyP(p)),
            br()
          ] : []),

          // CONCLUSÃO
          secNum(obsGeral ? 8 + numOffset : 7 + numOffset, 'CONCLUSÃO'),
          ...conclusaoTexto.map(p => bodyP(p)),
          br(),

          // ASSINATURA
          secNum(obsGeral ? 9 + numOffset : 8 + numOffset, 'ASSINATURA DO RESPONSÁVEL'),
          br(),
          new Paragraph({ children: [N('_______________________________', 20)], alignment: AlignmentType.CENTER, spacing: { after: 60 } }),
          new Paragraph({ children: [N('Gabriela Soares Borges', 20)], alignment: AlignmentType.CENTER, spacing: { after: 40 } }),
          new Paragraph({ children: [N('Engenheira Civil · CREA: 427760MG', 20)], alignment: AlignmentType.CENTER, spacing: { after: 40 } }),

        ]
      }
    ]
  });

  return Packer.toBuffer(doc);
}

// ── Login por senha única (protege dados de clientes) ──────
function parseCookies(cabecalho) {
  const cookies = {};
  (cabecalho || '').split(';').forEach(par => {
    const i = par.indexOf('=');
    if (i === -1) return;
    cookies[par.slice(0, i).trim()] = decodeURIComponent(par.slice(i + 1).trim());
  });
  return cookies;
}

function estaAutenticado(req) {
  if (!APP_PASSWORD) return true; // sem senha configurada, não bloqueia (comportamento de antes)
  const cookies = parseCookies(req.headers.cookie);
  return cookies['engcheck_auth'] === APP_PASSWORD;
}

function paginaLogin(erro) {
  return `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>EngCheck · Login</title>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Montserrat',sans-serif;background:#1A3C5E;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{background:#fff;border-radius:16px;padding:32px 28px;max-width:340px;width:100%;text-align:center;box-shadow:0 20px 50px rgba(0,0,0,.25)}
img{width:64px;height:64px;object-fit:contain;background:#fff;border-radius:12px;padding:6px;border:1px solid #eee;margin-bottom:16px}
h1{font-size:22px;font-weight:800;color:#1A3C5E;margin-bottom:4px}
p.sub{font-size:12px;color:#888;margin-bottom:22px}
input{width:100%;padding:13px 14px;border:1.5px solid #ddd;border-radius:9px;font-size:15px;font-family:inherit;margin-bottom:14px;outline:none}
input:focus{border-color:#D4762A}
button{width:100%;background:#D4762A;color:#fff;border:none;border-radius:9px;padding:13px;font-size:14px;font-weight:700;font-family:inherit;cursor:pointer}
.erro{background:#FDECEA;color:#C0392B;font-size:12px;font-weight:600;padding:9px;border-radius:8px;margin-bottom:14px}
</style></head>
<body>
  <form class="card" method="POST" action="/login">
    <img src="/logo.png" alt="EngCheck">
    <h1>EngCheck</h1>
    <p class="sub">Digite a senha de acesso</p>
    ${erro ? '<div class="erro">Senha incorreta. Tente de novo.</div>' : ''}
    <input type="password" name="senha" placeholder="Senha" autofocus required>
    <button type="submit">Entrar</button>
  </form>
</body></html>`;
}


// ── HTTP SERVER ──────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  const caminho = req.url.split('?')[0];

  // Login: mostra o formulário, ou confere a senha enviada
  if (caminho === '/login') {
    if (req.method === 'GET') {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      return res.end(paginaLogin(false));
    }
    if (req.method === 'POST') {
      let body = '';
      req.on('data', chunk => body += chunk);
      req.on('end', () => {
        const params = new URLSearchParams(body);
        const senha = params.get('senha') || '';
        if (APP_PASSWORD && senha === APP_PASSWORD) {
          res.writeHead(302, {
            'Set-Cookie': `engcheck_auth=${encodeURIComponent(senha)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=31536000`,
            Location: '/'
          });
          return res.end();
        }
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        return res.end(paginaLogin(true));
      });
      return;
    }
  }

  if (caminho === '/logout') {
    res.writeHead(302, { 'Set-Cookie': 'engcheck_auth=; Path=/; HttpOnly; Max-Age=0', Location: '/login' });
    return res.end();
  }

  // Bloqueia todo o resto do app se a senha estiver configurada e a pessoa não tiver logado
  // (ícones/manifest do PWA ficam liberados, senão a própria tela de login não carrega o logo)
  const PUBLICO = ['/logo.png', '/icon-192.png', '/icon-512.png', '/manifest.json', '/service-worker.js'];
  if (!PUBLICO.includes(caminho) && !estaAutenticado(req)) {
    if (req.method === 'GET' && (req.headers.accept || '').includes('text/html')) {
      res.writeHead(302, { Location: '/login' });
      return res.end();
    }
    res.writeHead(401, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ erro: 'Não autorizado. Faça login novamente.' }));
  }

  if (req.method === 'GET' && req.url === '/') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    return res.end(fs.readFileSync(path.join(__dirname, 'vistoria_app.html')));
  }
  if (req.method === 'GET' && req.url === '/logo.png') {
    res.writeHead(200, { 'Content-Type': 'image/png' });
    return res.end(fs.readFileSync(LOGO_IMG));
  }
  if (req.method === 'GET' && req.url === '/manifest.json') {
    res.writeHead(200, { 'Content-Type': 'application/manifest+json' });
    return res.end(fs.readFileSync(path.join(__dirname, 'manifest.json')));
  }
  if (req.method === 'GET' && req.url === '/service-worker.js') {
    // Service-Worker-Allowed garante que o service worker consegue controlar todo o site (escopo "/")
    res.writeHead(200, { 'Content-Type': 'application/javascript; charset=utf-8', 'Service-Worker-Allowed': '/' });
    return res.end(fs.readFileSync(path.join(__dirname, 'service-worker.js')));
  }
  if (req.method === 'GET' && (req.url === '/icon-192.png' || req.url === '/icon-512.png')) {
    res.writeHead(200, { 'Content-Type': 'image/png' });
    return res.end(fs.readFileSync(path.join(__dirname, req.url.slice(1))));
  }
  if (req.method === 'GET' && req.url.startsWith('/obter-mapa')) {
    try {
      const urlObj = new URL(req.url, `http://${req.headers.host}`);
      const lat = parseFloat(urlObj.searchParams.get('lat'));
      const lon = parseFloat(urlObj.searchParams.get('lon'));
      if (isNaN(lat) || isNaN(lon)) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ erro: 'Coordenadas inválidas.' }));
      }
      if (!GEOAPIFY_API_KEY) {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ erro: 'Chave do serviço de mapas (GEOAPIFY_API_KEY) não configurada no servidor.' }));
      }
      // Geoapify Static Maps: serviço confiável baseado no OpenStreetMap, com tier gratuito
      const mapaUrl = `https://maps.geoapify.com/v1/staticmap?style=osm-bright&width=640&height=400&center=lonlat:${lon},${lat}&zoom=13&marker=lonlat:${lon},${lat};type:awesome;color:orange;size:large&apiKey=${GEOAPIFY_API_KEY}`;
      const mapaRes = await fetch(mapaUrl);
      if (!mapaRes.ok) {
        const corpoErro = await mapaRes.text().catch(() => '');
        throw new Error(`Geoapify respondeu ${mapaRes.status}: ${corpoErro.slice(0, 200)}`);
      }
      const buf = Buffer.from(await mapaRes.arrayBuffer());
      const base64 = buf.toString('base64');
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ base64, mediaType: 'image/png' }));
    } catch (e) {
      console.error('Erro em /obter-mapa:', e.message);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ erro: `Não foi possível buscar o mapa automaticamente. Detalhe: ${e.message}` }));
    }
  }

  // Dispara a compactação na hora, sem esperar a rotina automática diária — útil pra testar.
  if (req.method === 'GET' && req.url === '/compactar-agora') {
    try {
      compactarLaudosAntigos();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ ok: true }));
    } catch (e) {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ ok: false, erro: e.message }));
    }
  }

  // Espaço usado no Volume do Railway (onde os laudos ficam guardados) — pra acompanhar
  // se está chegando perto do limite antes que o app pare de conseguir salvar coisa nova.
  if (req.method === 'GET' && req.url === '/armazenamento') {
    try {
      const saida = execSync(`df -k "${OUTPUT_DIR}"`).toString();
      const linhas = saida.trim().split('\n');
      const partes = linhas[linhas.length - 1].trim().split(/\s+/);
      const totalKB = parseInt(partes[1], 10);
      const usadoKB = parseInt(partes[2], 10);
      const percentual = (totalKB > 0) ? Math.round((usadoKB / totalKB) * 100) : null;
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ usadoMB: Math.round(usadoKB / 1024), totalMB: Math.round(totalKB / 1024), percentual }));
    } catch (e) {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ erro: e.message }));
    }
  }

  // ── Agenda: pré-cadastro dos dados do cliente/imóvel antes da vistoria ──
  if (req.method === 'GET' && req.url === '/agenda') {
    const lista = lerAgenda().sort((a, b) => (a.dados.data || '9999').localeCompare(b.dados.data || '9999'));
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify(lista));
  }

  if (req.method === 'POST' && req.url === '/agenda') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body);
        if (!payload.dados || !payload.dados.nome) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          return res.end(JSON.stringify({ erro: 'Nome do cliente é obrigatório.' }));
        }
        const agenda = lerAgenda();
        const item = {
          id: 'ag_' + Date.now() + Math.random().toString(36).slice(2),
          dados: payload.dados,
          plantaBase64: payload.plantaBase64 || null,
          plantaMediaType: payload.plantaMediaType || null,
          mapaBase64: payload.mapaBase64 || null,
          mapaMediaType: payload.mapaMediaType || null,
          criadoEm: Date.now()
        };
        agenda.push(item);
        salvarAgenda(agenda);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ ok: true, id: item.id }));
      } catch (e) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ erro: e.message }));
      }
    });
    return;
  }

  if (req.method === 'DELETE' && req.url.startsWith('/agenda')) {
    const urlObj = new URL(req.url, `http://${req.headers.host}`);
    const id = urlObj.searchParams.get('id');
    const agenda = lerAgenda().filter(a => a.id !== id);
    salvarAgenda(agenda);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ ok: true }));
  }

  if (req.method === 'PUT' && req.url.startsWith('/agenda')) {
    const urlObj = new URL(req.url, `http://${req.headers.host}`);
    const id = urlObj.searchParams.get('id');
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const payload = JSON.parse(body);
        const agenda = lerAgenda();
        const item = agenda.find(a => a.id === id);
        if (!item) {
          res.writeHead(404, { 'Content-Type': 'application/json' });
          return res.end(JSON.stringify({ erro: 'Agendamento não encontrado.' }));
        }
        item.dados = payload.dados;
        item.plantaBase64 = payload.plantaBase64 || null;
        item.plantaMediaType = payload.plantaMediaType || null;
        item.mapaBase64 = payload.mapaBase64 || null;
        item.mapaMediaType = payload.mapaMediaType || null;
        salvarAgenda(agenda);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ ok: true }));
      } catch (e) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ erro: e.message }));
      }
    });
    return;
  }

  if (req.method === 'GET' && req.url === '/historico') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify(lerHistorico().reverse()));
  }
  if (req.method === 'GET' && req.url.startsWith('/download-pdf/')) {
    const filename = decodeURIComponent(req.url.replace('/download-pdf/', ''));
    const docxPath = path.join(OUTPUT_DIR, filename);
    const pdfPath = docxPath.replace('.docx', '.pdf');
    const pdfBuf = lerArquivoTalvezCompactado(pdfPath);
    if (pdfBuf) {
      res.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `attachment; filename="${filename.replace('.docx','.pdf')}"` });
      return res.end(pdfBuf);
    }
    // Tenta gerar o PDF a partir do docx (se o docx existir, mesmo que compactado)
    const docxBuf = lerArquivoTalvezCompactado(docxPath);
    if (docxBuf) {
      try {
        if (!fs.existsSync(docxPath)) fs.writeFileSync(docxPath, docxBuf); // descompacta temporariamente pro LibreOffice conseguir ler
        const paths = ['C:\\Program Files\\LibreOffice\\program\\soffice.exe','C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe','soffice'];
        let soffice = 'soffice';
        for (const p of paths) { if (fs.existsSync(p)) { soffice = p; break; } }
        execSync(`"${soffice}" --headless --convert-to pdf "${docxPath}" --outdir "${OUTPUT_DIR}"`, { timeout: 60000 });
        const pdfBufGerado = fs.readFileSync(pdfPath);
        res.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `attachment; filename="${filename.replace('.docx','.pdf')}"` });
        return res.end(pdfBufGerado);
      } catch(e) {
        // fallback: manda o docx mesmo
        res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Content-Disposition': `attachment; filename="${filename}"` });
        return res.end(docxBuf);
      }
    }
    res.writeHead(404); return res.end('Not found');
  }

  if (req.method === 'GET' && req.url.startsWith('/download/')) {
    const filename = decodeURIComponent(req.url.replace('/download/', ''));
    const filepath = path.join(OUTPUT_DIR, filename);
    const buf = lerArquivoTalvezCompactado(filepath);
    if (buf) {
      res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Content-Disposition': `attachment; filename="${filename}"` });
      return res.end(buf);
    }
    res.writeHead(404); return res.end('Not found');
  }

  if (req.method === 'POST' && req.url === '/gerar-conclusao') {
    try {
      const { registros, tipoVistoria, obsGeral } = await lerCorpoJSON(req);
      const listaDefeitos = (registros || [])
        .map(r => `- ${r.ambiente}: ${r.defeito}`)
        .join('\n') || 'Nenhum defeito registrado.';
      const outrosProblemas = (obsGeral || '').trim();
      const texto = await chamarClaude({
        maxTokens: 700,
        system: 'Você é um(a) engenheiro(a) civil redigindo a seção de CONCLUSÃO de um laudo técnico de vistoria de imóvel, seguindo ABNT NBR 16747:2020, ABNT NBR 5674:2024 e conceitos do IBAPE Nacional. Escreva 3 a 4 parágrafos técnicos, objetivos e formais, em português, baseados nos defeitos e observações fornecidos, recomendando a correção antes da entrega/aceitação do imóvel. Retorne SOMENTE os parágrafos de texto, separados por uma linha em branco, sem títulos, sem markdown, sem numeração.',
        messages: [{ role: 'user', content: `Tipo de vistoria: ${tipoVistoria || 'não informado'}\n\nDefeitos registrados (por legenda escolhida em cada foto):\n${listaDefeitos}\n\nOutros problemas observados:\n${outrosProblemas || 'Nenhum'}\n\nRedija a conclusão do laudo.` }]
      });
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ conclusao: texto.trim() }));
    } catch (e) {
      console.error('Erro em /gerar-conclusao:', e.message);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ conclusao: '' })); // falha silenciosa: server usa o texto padrão fixo
    }
  }

  // Transforma uma descrição curta/informal de defeito (ex: "buraco na parede") no
  // texto técnico formal que vai pro laudo, no mesmo estilo da biblioteca de legendas.
  if (req.method === 'POST' && req.url === '/gerar-defeito-texto') {
    try {
      const { ambiente, textoCurto } = await lerCorpoJSON(req);
      if (!textoCurto || !textoCurto.trim()) throw new Error('Texto do defeito vazio.');
      const texto = await chamarClaude({
        maxTokens: 200,
        system: 'Você é um(a) engenheiro(a) civil redigindo a descrição técnica de UM defeito construtivo específico, pra entrar na seção de registros fotográficos de um laudo de vistoria de imóvel. Use o mesmo estilo formal e objetivo de normas ABNT e do IBAPE Nacional. Escreva de 1 a 2 frases apenas, descrevendo o defeito de forma técnica, e se fizer sentido, uma recomendação breve de correção. NÃO use saudações, títulos, aspas, markdown ou numeração — retorne SOMENTE o texto final, pronto pra ser colado no laudo, como neste exemplo de estilo: "Foram identificados danos na superfície da parede, comprometendo a integridade e o acabamento do revestimento. Recomenda-se reparo com massa adequada e repintura do trecho afetado."',
        messages: [{ role: 'user', content: `Ambiente: ${ambiente || 'não informado'}\nDefeito descrito de forma resumida pelo usuário: "${textoCurto.trim()}"\n\nRedija a descrição técnica formal desse defeito, pronta pra entrar no laudo.` }]
      });
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ texto: texto.trim() }));
    } catch (e) {
      console.error('Erro em /gerar-defeito-texto:', e.message);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ erro: e.message }));
    }
  }

  if (req.method === 'POST' && req.url === '/gerar-laudo') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const payload = JSON.parse(body);
        const docxBuffer = await gerarDocx(payload);
        const ts = Date.now();
        const nomeArquivo = `Laudo_${(payload.dados.nome||'cliente').replace(/\s+/g,'_')}_${payload.dados.data||'vistoria'}_${ts}.docx`;
        const docxPath = path.join(OUTPUT_DIR, nomeArquivo);
        fs.writeFileSync(docxPath, docxBuffer);

        const hist = lerHistorico();
        const registrosResumo = (payload.registros || []).map(r => ({ ambiente: r.ambiente, defeito: r.defeito }));
        hist.push({ nome: payload.dados.nome||'Cliente', empreendimento: payload.dados.empreendimento||'', endereco: payload.dados.endereco||'', data: payload.dados.data||'—', tipo: payload.tipoVistoria||'—', valor: payload.dados.valor || null, formaPagamento: payload.dados.formaPagamento || null, registros: registrosResumo, arquivo: nomeArquivo, ts });
        salvarHistorico(hist);

        if (payload.formato === 'docx') {
          res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Content-Disposition': `attachment; filename="${nomeArquivo}"` });
          return res.end(docxBuffer);
        }

        try {
          const paths = ['C:\\Program Files\\LibreOffice\\program\\soffice.exe','C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe','soffice'];
          let soffice = 'soffice';
          for (const p of paths) { if (fs.existsSync(p)) { soffice = p; break; } }
          execSync(`"${soffice}" --headless --convert-to pdf "${docxPath}" --outdir "${OUTPUT_DIR}"`, { timeout: 60000 });
          const pdfBuf = fs.readFileSync(docxPath.replace('.docx','.pdf'));
          res.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `attachment; filename="${nomeArquivo.replace('.docx','.pdf')}"` });
          return res.end(pdfBuf);
        } catch(e) {
          console.error('PDF falhou, enviando docx:', e.message);
          res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Content-Disposition': `attachment; filename="${nomeArquivo}"` });
          return res.end(docxBuffer);
        }
      } catch(e) {
        console.error('Erro:', e);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: e.message }));
      }
    });
    return;
  }

  res.writeHead(404); res.end('Not found');
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`\n✅ Servidor rodando em http://localhost:${PORT}`);
  console.log(`   Acesse no celular: descubra seu IP com "ipconfig" e use http://SEU_IP:${PORT}\n`);
});
