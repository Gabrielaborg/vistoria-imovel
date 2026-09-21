#!/usr/bin/env node
const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const OUTPUT_DIR = path.join(__dirname, 'laudos_gerados');
const COVER_IMG = path.join(__dirname, 'cover_page.png');
const FOOTER_IMG = path.join(__dirname, 'footer_bar.jpeg');
const LOGO_IMG = path.join(__dirname, 'logo.png');
const HISTORICO_FILE = path.join(OUTPUT_DIR, 'historico.json');

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

function lerHistorico() { try { return JSON.parse(fs.readFileSync(HISTORICO_FILE, 'utf8')); } catch { return []; } }
function salvarHistorico(h) { try { fs.writeFileSync(HISTORICO_FILE, JSON.stringify(h, null, 2)); } catch(e) { console.error('Erro histórico:', e.message); } }

async function gerarDocx(payload) {
  const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, Footer } = require('docx');
  const { dados, tipoVistoria, obsGeral, registros, plantaBase64, plantaMediaType, mapaBase64, mapaMediaType } = payload;

  const dataFmt = dados.data
    ? new Date(dados.data + 'T12:00:00').toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' })
    : new Date().toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' });

  const coverImgData = fs.readFileSync(COVER_IMG);
  const footerImgData = fs.readFileSync(FOOTER_IMG);

  const FONT = 'Montserrat';
  const B = (text, size=20) => new TextRun({ text, bold: true, size, font: FONT });
  const N = (text, size=20) => new TextRun({ text, size, font: FONT });
  const br = () => new Paragraph({ children: [N('')], spacing: { after: 120 } });
  const indent = { left: 720 };

  const secNum = (n, text) => new Paragraph({ children: [B(`${n}. ${text}`, 22)], spacing: { before: 300, after: 140 } });
  const subNum = (n, text) => new Paragraph({ children: [B(`${n} `, 20), B(text, 20)], indent, spacing: { before: 180, after: 80 } });
  const bodyP = (text) => new Paragraph({ children: [N(text, 20)], indent, spacing: { after: 100 }, alignment: AlignmentType.JUSTIFIED });
  const bodyBold = (label, val) => new Paragraph({ children: [B(label, 20), N(val, 20)], indent, spacing: { after: 80 } });
  const bulletP = (text) => new Paragraph({ children: [N(`• ${text}`, 20)], indent: { left: 1260 }, spacing: { after: 60 } });
  const defP = (label, text) => new Paragraph({ children: [B(label+': ', 20), N(text, 20)], indent, spacing: { after: 80 }, alignment: AlignmentType.JUSTIFIED });

  const makeFooter = () => new Footer({
    children: [new Paragraph({
      children: [new ImageRun({ data: footerImgData, transformation: { width: 520, height: 17 }, type: 'jpg' })],
      alignment: AlignmentType.CENTER, spacing: { before: 60 }
    })]
  });

  // ─── Registros (sem texto de IA por cômodo) ─────────────
  const registrosParagraphs = [];
  let imgCounter = 1;
  const byAmbiente = {};
  registros.forEach(r => { if (!byAmbiente[r.ambiente]) byAmbiente[r.ambiente] = []; byAmbiente[r.ambiente].push(r); });

  for (const [ambiente, items] of Object.entries(byAmbiente)) {
    // Nome do ambiente em negrito
    registrosParagraphs.push(new Paragraph({
      children: [B(ambiente, 20)], indent, spacing: { before: 200, after: 100 }
    }));

    for (const item of items) {
      for (const foto of item.fotos) {
        // Só: "Imagem N - tipo do defeito" + foto
        const desc = item.defeito;
        registrosParagraphs.push(new Paragraph({
          children: [N(`Imagem ${imgCounter} - ${desc}`, 20)], indent, spacing: { after: 80 }
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
        new Paragraph({ children: [new ImageRun({ data: plantaBuf, transformation: { width: 420, height: 300 }, type: plantaType })], alignment: AlignmentType.CENTER, spacing: { before: 100, after: 160 } }),
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
        new Paragraph({ children: [new ImageRun({ data: mapaBuf, transformation: { width: 420, height: 220 }, type: mapaType })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 120 } })
      );
    } catch(e) { console.error('Mapa error:', e.message); }
  } else {
    mapaParagraphs.push(new Paragraph({ children: [B('Localização:', 20), N(' ' + [dados.endereco, dados.bloco, dados.apto, dados.cidade].filter(Boolean).join(', '), 20)], indent, spacing: { after: 80 } }));
  }

  const comodosLines = (dados.comodos||'').split('\n').map(l=>l.trim()).filter(Boolean);
  const comodosItems = comodosLines.length > 0 ? comodosLines.map(l => bulletP(l.replace(/[;,.]$/,'') + ';')) : [bulletP('Conforme especificações do projeto.')];
  const endFull = [dados.endereco, dados.bloco, dados.apto, dados.cidade, dados.cep].filter(Boolean).join(', ');
  const numOffset = plantaParagraphs.length > 0 ? 1 : 0;

  // ─── Conclusão robusta ───────────────────────────────────
  const conclusaoTexto = [
    `Com base na vistoria técnica realizada na unidade habitacional em ${dataFmt}, constatou-se a presença de não conformidades construtivas, falhas de acabamento e inconformidades funcionais distribuídas nos ambientes inspecionados, conforme amplamente documentado por meio de registros fotográficos e descrições técnicas ao longo deste relatório. As manifestações identificadas incluem irregularidades em revestimentos cerâmicos, falhas de rejuntamento, defeitos em pintura, problemas em esquadrias, portas, elementos hidráulicos e demais sistemas construtivos aparentes.`,
    `As não conformidades constatadas não se limitam a questões meramente estéticas. Diversas das anomalias identificadas têm potencial para comprometer diretamente a durabilidade dos materiais, a estanqueidade da edificação, o desempenho dos sistemas construtivos e a habitabilidade do imóvel ao longo do tempo. Falhas no rejunte e revestimentos com som cavo, por exemplo, podem evoluir para infiltrações, proliferação de fungos e desplacamento de peças, representando risco à segurança dos moradores e gerando custos elevados de reparo futuro. Fissuras e trincas em paredes, quando não tratadas, podem indicar movimentações estruturais ou de revestimento que tendem a se agravar progressivamente.`,
    `Do ponto de vista jurídico e técnico, a construtora/incorporadora é responsável pela qualidade de execução da obra e pela entrega do imóvel em conformidade com os padrões estabelecidos pela ABNT NBR 15575 – Desempenho de Edificações Habitacionais, bem como pelas normas de execução aplicáveis a cada sistema construtivo. Os defeitos registrados neste laudo evidenciam falhas no controle de qualidade durante a execução e na fase de entrega, sendo tecnicamente incompatíveis com o padrão esperado para um imóvel novo. Nos termos do Código Civil Brasileiro (art. 618) e do Código de Defesa do Consumidor (art. 26), o proprietário tem direito à reparação integral dos vícios construtivos identificados, dentro dos prazos legais de garantia.`,
    `Portanto, conclui-se que todas as não conformidades registradas neste documento devem ser integralmente corrigidas pela construtora/responsável técnico, sem ônus ao proprietário, antes da aceitação definitiva do imóvel. Recomenda-se formalmente que: (1) o presente laudo seja apresentado à construtora como instrumento técnico de notificação; (2) os reparos sejam executados com acompanhamento técnico qualificado e observância às normas vigentes; (3) após a execução dos reparos, seja realizada nova vistoria de conferência para atestar a conformidade das correções; e (4) o proprietário não assine o Termo de Entrega de Chaves sem a prévia resolução das pendências aqui apontadas, preservando assim seus direitos legais de garantia.`
  ];

  const doc = new Document({
    sections: [
      {
        properties: { page: { size: { width: 11910, height: 16840 }, margin: { top: 0, right: 0, bottom: 0, left: 0 } } },
        children: [new Paragraph({ children: [new ImageRun({ data: coverImgData, transformation: { width: 794, height: 1123 }, type: 'png' })], alignment: AlignmentType.CENTER })]
      },
      {
        properties: { page: { size: { width: 11910, height: 16840 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
        footers: { default: makeFooter() },
        children: [
          secNum(1, 'IDENTIFICAÇÃO DO CONTRATANTE'),
          bodyP(`${dados.nome || 'Cliente'}, portador(a) do CPF nº ${dados.cpf || '—'}${dados.telefone ? `, telefone ${dados.telefone}` : ''}, solicitou a elaboração do presente relatório de vistoria de imóvel, com o objetivo de registrar as condições da unidade no momento da entrega, identificando eventuais inconformidades aparentes e falhas de execução visíveis.`),
          br(),
          secNum(2, 'OBJETIVO'),
          bodyP(`Este relatório tem como finalidade documentar, de forma objetiva e detalhada, as condições do imóvel ${tipoVistoria === 'Imóvel Novo' ? 'novo ' : ''}na data da vistoria, identificando eventuais não conformidades em acabamentos, instalações elétricas e hidráulicas, estrutura e funcionalidade dos ambientes. A avaliação foi conduzida seguindo as diretrizes estabelecidas pelas normas técnicas vigentes, incluindo as NBR (Normas Brasileiras) e os referenciais do PBQP-H, assegurando que os padrões de qualidade, segurança, funcionalidade e durabilidade do empreendimento sejam observados.`),
          br(),
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
          secNum(4, 'DESCRIÇÃO DO IMÓVEL'),
          bodyP('O imóvel vistoriado trata-se de um apartamento com a seguinte configuração:'),
          ...(dados.metragem ? [bulletP(`Área total: ${dados.metragem} m²`)] : []),
          ...comodosItems,
          br(),
          bodyP('Durante a vistoria, foram inspecionados os acabamentos, instalações elétricas e hidráulicas, funcionalidade dos ambientes e demais itens que compõem o imóvel, registrando-se eventuais não conformidades para que sejam corrigidas conforme os padrões de qualidade estabelecidos pela construtora.'),
          br(),
          ...plantaParagraphs,
          secNum(5 + numOffset, 'ELABORAÇÃO DE RELATÓRIO'),
          bodyP('A elaboração do presente relatório de vistoria técnica de recebimento da unidade habitacional foi realizada com base na identificação dos elementos construtivos aparentes, sua localização dentro do imóvel e as manifestações patológicas visíveis no momento da inspeção.'),
          br(),
          bodyP('Durante a vistoria, foram observados diversos pontos de não conformidade, falhas de acabamento, anomalias e possíveis vícios construtivos que podem comprometer o desempenho esperado dos sistemas e materiais.'),
          br(),
          bodyP('A inspeção foi feita com base nos princípios estabelecidos pela ABNT NBR 16747:2020 – Diretrizes para inspeção predial, na ABNT NBR 5674:2024 – Manutenção de edificações, e também conforme os conceitos definidos pelo IBAPE Nacional.'),
          br(),
          bodyP('Considerando que alguns termos utilizados neste documento podem não ser de conhecimento geral, seguem abaixo os principais conceitos utilizados ao longo do relatório:'),
          br(),
          defP('Anomalia', 'Irregularidade que compromete o desempenho de um elemento ou sistema da edificação. Pode ter origem no projeto, execução, uso ou manutenção inadequada.'),
          defP('Manifestação Patológica', 'Sinais visíveis de degradação, como fissuras, manchas, destacamentos, entre outros.'),
          defP('Agente de Degradação', 'Fatores (naturais, físicos ou químicos) que contribuem para a deterioração dos elementos construtivos.'),
          defP('Falha', 'Perda da função de um componente, seja por uso indevido, má execução ou falta de manutenção.'),
          defP('Desempenho', 'Comportamento da edificação e seus sistemas durante o uso, frente às solicitações normais esperadas ao longo de sua vida útil.'),
          defP('Vida Útil (VU)', 'Período em que um sistema ou componente deve cumprir suas funções, conforme previsto em projeto e respeitada sua manutenção adequada.'),
          defP('Plano de Manutenção', 'Documento técnico que organiza as ações necessárias de manutenção preventiva e corretiva de uma edificação.'),
          br(),
          bodyP('A unidade inspecionada apresenta diversas não conformidades visuais. Tais ocorrências indicam ausência de cuidados na execução final e comprometem o recebimento do imóvel em condições ideais de entrega.'),
          br(),
          bodyP('A recomendação técnica é que todas as anomalias listadas neste relatório sejam corrigidas antes da conclusão da entrega da unidade ao proprietário, garantindo o desempenho mínimo esperado e evitando prejuízos futuros.'),
          br(),
          secNum(6 + numOffset, 'REGISTRO DE NÃO CONFORMIDADES DA VISTORIA'),
          bodyP(`A seguir, são apresentados os registros fotográficos das não conformidades identificadas durante a vistoria no dia ${dataFmt}, acompanhados da respectiva descrição detalhada.`),
          br(),
          ...registrosParagraphs,
          ...(obsGeral ? [secNum(7 + numOffset, 'OUTROS PROBLEMAS'), ...obsGeral.split('\n').map(p => bodyP(p)), br()] : []),
          secNum(obsGeral ? 8 + numOffset : 7 + numOffset, 'CONCLUSÃO'),
          ...conclusaoTexto.map(p => bodyP(p)),
          br(),
          secNum(obsGeral ? 9 + numOffset : 8 + numOffset, 'ASSINATURA DO RESPONSÁVEL'),
          br(),
          new Paragraph({ children: [N('_______________________________', 20)], alignment: AlignmentType.CENTER, spacing: { after: 60 } }),
          new Paragraph({ children: [N('Gabriela Soares Borges', 20)], alignment: AlignmentType.CENTER, spacing: { after: 40 } }),
          new Paragraph({ children: [N('Engenheira Civil · CREA: 427760MG', 20)], alignment: AlignmentType.CENTER }),
        ]
      }
    ]
  });

  return Packer.toBuffer(doc);
}

const server = http.createServer(async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');

  if (req.method === 'GET' && req.url === '/') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    return res.end(fs.readFileSync(path.join(__dirname, 'vistoria_app.html')));
  }
  if (req.method === 'GET' && req.url === '/logo.png') {
    res.writeHead(200, { 'Content-Type': 'image/png' });
    return res.end(fs.readFileSync(LOGO_IMG));
  }
  if (req.method === 'GET' && req.url === '/historico') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify(lerHistorico().reverse()));
  }
  if (req.method === 'GET' && req.url.startsWith('/download/')) {
    const filename = decodeURIComponent(req.url.replace('/download/', ''));
    const filepath = path.join(OUTPUT_DIR, filename);
    if (fs.existsSync(filepath)) {
      res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Content-Disposition': `attachment; filename="${filename}"` });
      return res.end(fs.readFileSync(filepath));
    }
    res.writeHead(404); return res.end('Not found');
  }

  if (req.method === 'POST' && req.url === '/gerar-laudo') {
    let body = '';
    let bodySize = 0;
    const MAX_SIZE = 50 * 1024 * 1024;
    req.on('data', chunk => { bodySize += chunk.length; if (bodySize <= MAX_SIZE) body += chunk; });
    req.on('end', async () => {
      try {
        if (bodySize > MAX_SIZE) throw new Error('Payload muito grande. Reduza o número ou tamanho das fotos.');
        const payload = JSON.parse(body);
        console.log(`Gerando: ${payload.dados?.nome} - ${payload.registros?.length} registros`);
        const docxBuffer = await gerarDocx(payload);
        const ts = Date.now();
        const nomeArquivo = `Laudo_${(payload.dados.nome||'cliente').replace(/\s+/g,'_')}_${payload.dados.data||'vistoria'}_${ts}.docx`;
        const docxPath = path.join(OUTPUT_DIR, nomeArquivo);
        fs.writeFileSync(docxPath, docxBuffer);
        const hist = lerHistorico();
        hist.push({ nome: payload.dados.nome||'Cliente', empreendimento: payload.dados.empreendimento||'', data: payload.dados.data||'—', tipo: payload.tipoVistoria||'—', arquivo: nomeArquivo, ts });
        salvarHistorico(hist);
        console.log(`Laudo gerado: ${nomeArquivo}`);
        res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Content-Disposition': `attachment; filename="${nomeArquivo}"` });
        return res.end(docxBuffer);
      } catch(e) {
        console.error('Erro:', e.message);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: e.message }));
      }
    });
    return;
  }

  res.writeHead(404); res.end('Not found');
});

server.listen(PORT, '0.0.0.0', () => console.log(`✅ Servidor na porta ${PORT}`));
