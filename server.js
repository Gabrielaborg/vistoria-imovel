#!/usr/bin/env node
const http = require('http');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const PORT = 3000;
const OUTPUT_DIR = path.join(__dirname, 'laudos_gerados');
const COVER_IMG = path.join(__dirname, 'cover_page.png');
const FOOTER_IMG = path.join(__dirname, 'footer_bar.jpeg');
const LOGO_IMG = path.join(__dirname, 'logo.png');
const HISTORICO_FILE = path.join(OUTPUT_DIR, 'historico.json');

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

function lerHistorico() { try { return JSON.parse(fs.readFileSync(HISTORICO_FILE, 'utf8')); } catch { return []; } }
function salvarHistorico(h) { fs.writeFileSync(HISTORICO_FILE, JSON.stringify(h, null, 2)); }

async function gerarDocx(payload) {
  const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, Footer } = require('docx');
  const { dados, tipoVistoria, obsGeral, registros, plantaBase64, plantaMediaType, mapaBase64, mapaMediaType } = payload;

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
  let imgCounter = 1;
  const byAmbiente = {};
  registros.forEach(r => { if (!byAmbiente[r.ambiente]) byAmbiente[r.ambiente] = []; byAmbiente[r.ambiente].push(r); });

  for (const [ambiente, items] of Object.entries(byAmbiente)) {
    registrosParagraphs.push(new Paragraph({
      children: [B(ambiente, 20)], indent, spacing: { before: 200, after: 100 }
    }));
    for (const item of items) {
      for (const foto of item.fotos) {
        const desc = foto.aiDesc || item.defeito;
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

  const conclusaoTexto = [
    'Com base na vistoria técnica realizada na unidade habitacional, constatou-se a presença de não conformidades construtivas, falhas de acabamento e inconformidades funcionais distribuídas nos ambientes inspecionados, conforme descrito e documentado ao longo deste relatório técnico. As manifestações observadas incluem irregularidades em revestimentos, falhas de rejuntamento, defeitos em pintura, problemas em esquadrias, portas, elementos hidráulicos, acabamentos e demais sistemas construtivos aparentes.',
    'Os defeitos identificados evidenciam deficiência nos processos executivos e no controle de qualidade durante as etapas de acabamento e entrega da unidade, não sendo compatíveis com o padrão esperado para um imóvel novo. Ainda que parte das inconformidades apresente caráter predominantemente estético, diversas manifestações podem comprometer a durabilidade dos materiais, o desempenho dos sistemas construtivos, a estanqueidade, a funcionalidade dos ambientes e a vida útil da edificação ao longo do tempo.',
    'Conforme os princípios estabelecidos pela ABNT NBR 15575, a edificação deve atender aos requisitos mínimos de desempenho relacionados à segurança, habitabilidade, funcionalidade e durabilidade. Da mesma forma, os serviços executivos e acabamentos devem seguir padrões adequados de qualidade e conformidade técnica, observando as boas práticas construtivas e as normas aplicáveis a cada sistema construtivo. As anomalias constatadas neste relatório demonstram inconformidades em relação a tais requisitos, tornando tecnicamente recomendável a correção integral dos itens apontados.',
    'Dessa forma, conclui-se que todas as não conformidades registradas neste documento devem ser devidamente corrigidas pela construtora/responsável técnico antes da aceitação definitiva do imóvel, garantindo o adequado desempenho dos sistemas, a preservação da vida útil dos materiais e o padrão de qualidade esperado para a edificação. Recomenda-se ainda que os reparos sejam executados com acompanhamento técnico e observância aos procedimentos normativos aplicáveis, a fim de evitar recorrência das falhas identificadas.'
  ];

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

// ── HTTP SERVER ──────────────────────────────────────────
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
  if (req.method === 'GET' && req.url.startsWith('/download-pdf/')) {
    const filename = decodeURIComponent(req.url.replace('/download-pdf/', ''));
    const docxPath = path.join(OUTPUT_DIR, filename);
    const pdfPath = docxPath.replace('.docx', '.pdf');
    if (fs.existsSync(pdfPath)) {
      res.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `attachment; filename="${filename.replace('.docx','.pdf')}"` });
      return res.end(fs.readFileSync(pdfPath));
    }
    // Try to generate PDF from docx
    if (fs.existsSync(docxPath)) {
      try {
        const paths = ['C:\\Program Files\\LibreOffice\\program\\soffice.exe','C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe','soffice'];
        let soffice = 'soffice';
        for (const p of paths) { if (fs.existsSync(p)) { soffice = p; break; } }
        execSync(`"${soffice}" --headless --convert-to pdf "${docxPath}" --outdir "${OUTPUT_DIR}"`, { timeout: 60000 });
        const pdfBuf = fs.readFileSync(pdfPath);
        res.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `attachment; filename="${filename.replace('.docx','.pdf')}"` });
        return res.end(pdfBuf);
      } catch(e) {
        // fallback: send docx
        res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Content-Disposition': `attachment; filename="${filename}"` });
        return res.end(fs.readFileSync(docxPath));
      }
    }
    res.writeHead(404); return res.end('Not found');
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
        hist.push({ nome: payload.dados.nome||'Cliente', empreendimento: payload.dados.empreendimento||'', data: payload.dados.data||'—', tipo: payload.tipoVistoria||'—', arquivo: nomeArquivo, ts });
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
