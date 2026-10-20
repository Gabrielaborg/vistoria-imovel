#!/usr/bin/env node
const http = require('http');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const { execSync } = require('child_process');

const PORT = 3000;
const OUTPUT_DIR = path.join(__dirname, 'laudos_gerados');
const COVER_IMG = path.join(__dirname, 'cover_page.png');
const COVER_CONSTATACAO_IMG = path.join(__dirname, 'cover_constatacao.jpg'); // capa "Laudo de Constatação"
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

// ── Vistorias não finalizadas (rascunhos) sincronizadas entre aparelhos ──
// Cada aparelho continua guardando suas vistorias localmente (funciona sem
// internet), e também envia uma cópia pra cá, assim todas aparecem em
// qualquer aparelho. Estrutura no Volume:
//   rascunhos/index.json              → resumo de cada vistoria (nome, tipo, versão...)
//   rascunhos/<id>/estado.json        → dados da vistoria, SEM as imagens
//   rascunhos/<id>/fotos/<fotoId>.json → cada foto, enviada uma única vez
// Vistorias finalizadas/apagadas viram uma "lápide" no índice, pra que uma
// cópia antiga em outro aparelho não faça a vistoria reaparecer sozinha.
const RASC_DIR = path.join(OUTPUT_DIR, 'rascunhos');
const RASC_INDEX = path.join(RASC_DIR, 'index.json');
if (!fs.existsSync(RASC_DIR)) fs.mkdirSync(RASC_DIR, { recursive: true });
const idValido = (id) => typeof id === 'string' && /^[A-Za-z0-9_]{1,100}$/.test(id);
function lerIndiceRascunhos() { try { return JSON.parse(fs.readFileSync(RASC_INDEX, 'utf8')); } catch { return {}; } }
function salvarIndiceRascunhos(idx) {
  const tmp = RASC_INDEX + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(idx));
  fs.renameSync(tmp, RASC_INDEX); // grava de forma atômica: nunca deixa o índice pela metade
}
const versaoRascunho = (r) => Math.max(r.atualizadoEm || 0, r.excluidoEm || 0, r.restauradoEm || 0);
const pastaRascunho = (id) => path.join(RASC_DIR, id);
const pastaFotosRascunho = (id) => path.join(RASC_DIR, id, 'fotos');
function idsFotosDoEstado(st) {
  const ids = [];
  (st.registros || []).forEach(r => (r.fotos || []).forEach(f => f && f.id && ids.push(f.id)));
  (st.fotosAtuais || []).forEach(f => f && f.id && ids.push(f.id));
  return ids;
}
function apagarPastaRascunho(id) { try { fs.rmSync(pastaRascunho(id), { recursive: true, force: true }); } catch (e) { console.error('Falha ao apagar rascunho', id, e.message); } }

// Apaga de vez vistorias não finalizadas que estão na lixeira há mais de 30 dias,
// e esquece lápides com mais de 120 dias.
function limparRascunhosAntigos() {
  const idx = lerIndiceRascunhos();
  const agora = Date.now();
  let mudou = false;
  for (const [id, m] of Object.entries(idx)) {
    if (m.removidoEm && agora - m.removidoEm > 120 * 24 * 60 * 60 * 1000) { delete idx[id]; mudou = true; continue; }
    if (!m.removidoEm && m.excluidoEm && agora - m.excluidoEm > 30 * 24 * 60 * 60 * 1000) {
      apagarPastaRascunho(id);
      idx[id] = { id, contaId: m.contaId, removidoEm: agora };
      mudou = true;
    }
  }
  if (mudou) salvarIndiceRascunhos(idx);
}
setTimeout(() => { try { limparRascunhosAntigos(); } catch (e) { console.error('Erro na limpeza de rascunhos:', e.message); } }, 60 * 1000);
setInterval(() => { try { limparRascunhosAntigos(); } catch (e) { console.error('Erro na limpeza de rascunhos:', e.message); } }, 24 * 60 * 60 * 1000);

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

// Roda periodicamente: apaga DE VEZ (arquivo + registro) os laudos que estão
// na lixeira há mais de 30 dias. Antes disso, o laudo excluído fica guardado
// como segunda chance — só depois desse prazo é removido sem volta.
const TRINTA_DIAS_MS = 30 * 24 * 60 * 60 * 1000;
function limparLixeiraAntiga() {
  const hist = lerHistorico();
  const agora = Date.now();
  const permanece = [];
  let removidos = 0;
  for (const item of hist) {
    if (item.excluidoEm && (agora - item.excluidoEm) >= TRINTA_DIAS_MS) {
      const docxPath = path.join(OUTPUT_DIR, item.arquivo);
      const pdfPath = docxPath.replace('.docx', '.pdf');
      [docxPath, docxPath + '.gz', pdfPath, pdfPath + '.gz'].forEach(p => {
        try { if (fs.existsSync(p)) fs.unlinkSync(p); } catch (e) { /* segue mesmo se algum arquivo já não existir */ }
      });
      removidos++;
      console.log(`✓ Laudo removido definitivamente da lixeira (30 dias): ${item.arquivo}`);
    } else {
      permanece.push(item);
    }
  }
  if (removidos > 0) salvarHistorico(permanece);
}

setTimeout(() => { try { limparLixeiraAntiga(); } catch (e) { console.error('Erro na rotina de limpeza da lixeira:', e.message); } }, 45 * 1000);
setInterval(() => { try { limparLixeiraAntiga(); } catch (e) { console.error('Erro na rotina de limpeza da lixeira:', e.message); } }, 24 * 60 * 60 * 1000);

// ── Chamada central à API da Anthropic (roda só no servidor, chave nunca exposta) ──
async function chamarClaude({ system, messages, maxTokens = 300, model = 'claude-sonnet-5' }) {
  if (!ANTHROPIC_API_KEY) throw new Error('ANTHROPIC_API_KEY não configurada no servidor.');
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model,
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
  // Importante: NÃO presumir que o primeiro bloco da resposta é sempre o texto —
  // o Claude às vezes manda um bloco de "raciocínio interno" antes do texto real,
  // e pegar só o content[0] nesses casos pegava um bloco vazio/errado, fazendo a
  // IA "parecer" que não respondeu nada. Filtra especificamente os blocos de texto.
  const blocosDeTexto = (data.content || []).filter(b => b.type === 'text').map(b => b.text);
  return blocosDeTexto.join('\n').trim();
}

// Para textos longos e técnicos (conclusão do Laudo de Constatação, anamnese), usa
// um modelo mais capaz; se ele não estiver disponível na conta, cai pro Haiku, que
// já é usado no resto do app — assim a geração nunca trava por causa do modelo.
// Dá pra trocar o modelo principal em Railway: Service > Variables > MODELO_LAUDO_TECNICO
const MODELO_LAUDO_TECNICO = process.env.MODELO_LAUDO_TECNICO || 'claude-sonnet-4-5';
async function chamarClaudeTecnico(opcoes) {
  const modelos = [MODELO_LAUDO_TECNICO, 'claude-haiku-4-5-20251001'];
  let ultimoErro = null;
  for (const model of [...new Set(modelos)]) {
    try { return await chamarClaude({ ...opcoes, model }); }
    catch (e) { ultimoErro = e; console.error(`Modelo ${model} falhou, tentando o próximo:`, e.message); }
  }
  throw ultimoErro;
}

// Lista de normas CONFERIDAS (número e assunto verificados em out/2026). A IA só
// pode citar normas desta lista — evita número de norma errado no laudo.
const NORMAS_REFERENCIA = 'ABNT NBR 15575 (desempenho de edificações habitacionais, partes 1 a 6); NBR 16747:2020 (inspeção predial); NBR 5674:2024 (manutenção); NBR 17170:2022 (garantias de edificações — prazos recomendados); NBR 13753 (revestimento de PISO interno ou externo com placas cerâmicas — use para qualquer defeito em piso, rodapé ou soleira); NBR 13754 (revestimento de PAREDES INTERNAS com placas cerâmicas — use para azulejo/porcelanato em paredes de banheiro, cozinha, área de serviço etc.); NBR 13755 (revestimento de PAREDES EXTERNAS E FACHADAS com placas cerâmicas — use SOMENTE para fachada ou parede externa, NUNCA para piso ou parede interna); NBR 13749 (revestimento de argamassa); NBR 13245 (execução de pinturas em edificações); NBR 9575 e 9574 (impermeabilização); NBR 5626 (água fria e quente); NBR 8160 (esgoto sanitário); NBR 10844 (águas pluviais — calhas e condutores); NBR 15575-5 (coberturas, inclusive estanqueidade, rufos e arremates); NBR 5410 (instalações elétricas); NBR 10821 (esquadrias); NBR 15930 (portas de madeira); NBR 6118 (projeto de estruturas de concreto); NBR 14931 (execução de estruturas de concreto); NBR 7190 (estruturas de madeira); NBR 16868 (alvenaria estrutural); NBR 6120 (ações/cargas para cálculo de estruturas, ex.: sobrecarga de reservatório); NBR 9050 (acessibilidade, quando pertinente).';
const REGRA_NORMAS = `REGRA OBRIGATÓRIA SOBRE NORMAS: cite SOMENTE normas desta lista conferida, com o número exatamente como aparece aqui: ${NORMAS_REFERENCIA} NÃO cite nenhuma outra norma, NÃO cite número de item, seção, tabela ou anexo, NÃO informe valores numéricos de tolerância atribuídos a norma e NÃO coloque ano em norma que não tenha ano nesta lista. Use a norma apenas para o assunto descrito entre parênteses — atenção especial aos revestimentos cerâmicos: PISO = NBR 13753; PAREDE INTERNA = NBR 13754; FACHADA/PAREDE EXTERNA = NBR 13755. Na dúvida, prefira "boas práticas de execução" a citar uma norma.`;

// Lê o corpo (body) de uma requisição POST como JSON
function lerCorpoJSON(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      try { resolve(JSON.parse(Buffer.concat(chunks).toString('utf8'))); } catch (e) { reject(e); }
    });
    req.on('error', reject);
  });
}

// Gera, por ambiente: (1) um parágrafo único consolidando todos os defeitos daquele
// cômodo numa narrativa técnica fluida (citando normas ABNT quando pertinente), e
// (2) uma legenda curta pra cada grupo de defeito, usada embaixo de cada foto —
// no estilo de laudo mais organizado (parágrafo geral + legendas objetivas por imagem).
// Gera, numa ÚNICA chamada de IA (independente de quantos ambientes a vistoria
// tenha), o parágrafo consolidado + legendas curtas de TODOS os ambientes de
// uma vez. Antes disso era uma chamada por ambiente — com uma vistoria de 8
// cômodos, eram 8 chamadas separadas só pra essa parte. Consolidar tudo numa
// chamada só, usando o modelo mais econômico (Haiku, que já dá conta bem desse
// tipo de tarefa de organizar texto), reduz bastante o custo por laudo.
// Modalidades da vistoria de Locação (escolhidas no app, dentro do botão Locação).
// O texto de "objetivo" entra na seção OBJETIVO do laudo; "contextoIA" orienta a
// IA ao escrever os parágrafos dos ambientes e a conclusão.
const MODALIDADES_LOCACAO = {
  'Entrada': {
    rotulo: 'vistoria de entrada de imóvel em locação',
    objetivo: 'Trata-se de VISTORIA DE ENTRADA, realizada no início da relação locatícia, previamente à ocupação do imóvel pelo locatário, com a finalidade de registrar o estado de conservação, funcionamento e acabamento em que o imóvel é entregue. Este registro servirá de referência para a comparação com as condições a serem verificadas na vistoria de saída, ao término da locação, nos termos do art. 22, inciso V, da Lei nº 8.245/1991, que assegura ao locatário a descrição minuciosa do estado do imóvel quando de sua entrega.',
    contextoIA: 'Esta é uma VISTORIA DE ENTRADA (início da locação, antes da ocupação pelo locatário): o objetivo é registrar o estado em que o imóvel está sendo entregue, para servir de referência futura. Nunca sugira que os danos foram causados pelo locatário.'
  },
  'Saída': {
    rotulo: 'vistoria de saída de imóvel em locação',
    objetivo: 'Trata-se de VISTORIA DE SAÍDA, realizada ao término da relação locatícia, após a desocupação do imóvel, com a finalidade de registrar o estado de conservação, funcionamento e acabamento em que o imóvel é restituído, possibilitando sua comparação com as condições registradas na vistoria de entrada. Nos termos do art. 23, inciso III, da Lei nº 8.245/1991, cabe ao locatário restituir o imóvel no estado em que o recebeu, salvo as deteriorações decorrentes do seu uso normal.',
    contextoIA: 'Esta é uma VISTORIA DE SAÍDA (fim da locação, após a desocupação): o objetivo é registrar o estado em que o imóvel está sendo restituído, para comparação com a vistoria de entrada. Diferencie com cuidado o que é compatível com desgaste natural do que é avaria, sem atribuir culpa de forma categórica — a comparação com a entrada é que define responsabilidades.'
  },
  'Levantamento': {
    rotulo: 'vistoria de levantamento do estado do imóvel',
    semComparacao: true,
    objetivo: 'Trata-se de VISTORIA DE LEVANTAMENTO, realizada sem caracterizar entrada ou saída de locatário, com a finalidade de registrar o estado atual de conservação, funcionamento e acabamento do imóvel e dos itens nele existentes na data da inspeção, permitindo ao(à) contratante conhecer as condições em que o imóvel se encontra no momento — seja durante a ocupação, seja para fins de acompanhamento, manutenção ou preparação de uma futura locação.',
    contextoIA: 'Esta é uma VISTORIA DE LEVANTAMENTO (não é entrada nem saída): o objetivo é registrar o estado atual do imóvel e do que existe nele, para conhecimento do(a) contratante. Não fale em comparação entre entrada e saída; destaque o que necessita manutenção.'
  }
};

async function gerarResumosDosAmbientes(listaAmbientes, tipoVistoria, contextoExtra = '') {
  const ehLocacao = tipoVistoria === 'Locação';
  const ehConstatacao = tipoVistoria === 'Constatação';
  // listaAmbientes: [{ ambiente, defeitos, inicioImg, fimImg, classificacaoPorDefeito }, ...]
  const comDefeitos = listaAmbientes.filter(a => a.defeitos && a.defeitos.length > 0);
  const resultado = {};
  listaAmbientes.forEach((a, i) => {
    if (!a.defeitos || a.defeitos.length === 0) {
      resultado[i] = { paragrafo: `Não foram identificadas não conformidades no ambiente ${a.ambiente} durante a vistoria.`, legendas: [] };
    }
  });
  if (comDefeitos.length === 0) return resultado;

  const blocos = comDefeitos.map((a, idx) => {
    const listaTexto = a.defeitos.map((d, i) => {
      const classi = ehLocacao && a.classificacaoPorDefeito && a.classificacaoPorDefeito[d];
      return `  ${i + 1}. ${d}${classi ? ` [Classificação informada: ${classi}]` : ''}`;
    }).join('\n');
    const faixaTexto = a.fimImg > a.inicioImg ? `nº ${a.inicioImg} a ${a.fimImg}` : `nº ${a.inicioImg}`;
    return `### Ambiente ${idx} — ${a.ambiente} (imagens ${faixaTexto})\n${listaTexto}`;
  }).join('\n\n');

  const system = ehConstatacao
    ? 'Você é um(a) engenheiro(a) civil redigindo as seções de ambientes de um LAUDO DE CONSTATAÇÃO de imóvel (registro técnico, por inspeção exclusivamente visual, das manifestações patológicas, danos e não conformidades aparentes), seguindo a ABNT NBR 16747:2020, a ABNT NBR 15575 e normas correlatas da lista conferida informada abaixo, além dos conceitos do IBAPE Nacional. Você vai receber VÁRIOS ambientes de uma vez, cada um com suas constatações numeradas. Responda SOMENTE com um JSON válido, sem markdown, sem texto antes ou depois, no formato exato: {"ambientes": [{"paragrafo": "...", "legendas": ["...", "..."]}, ...]}. O array "ambientes" deve ter EXATAMENTE um item por ambiente informado, NA MESMA ORDEM em que foram apresentados. Em cada item: "paragrafo" é um único parágrafo fluido e técnico (não uma lista, não repita os textos originais colados um atrás do outro) descrevendo TODAS as manifestações constatadas naquele ambiente, citando a norma ABNT mais pertinente quando fizer sentido, usando linguagem cautelosa quanto à origem (ex: "compatível com", "indícios de") sem afirmar causa ou responsável de forma categórica, e mencionando que as constatações estão registradas nas imagens indicadas daquele ambiente, que integram o laudo. "legendas" deve ter EXATAMENTE um item por constatação daquele ambiente, na mesma ordem, cada uma uma legenda curta (até 10 palavras) descrevendo objetivamente o que a foto mostra (ex: "Marcas de infiltração por capilaridade e bolor").'
    : ehLocacao
    ? 'Você é um(a) engenheiro(a) civil redigindo as seções de ambientes de um laudo técnico de VISTORIA DE LOCAÇÃO de imóvel (não é vistoria de entrega de obra nova), considerando as disposições da Lei nº 8.245/1991 (Lei do Inquilinato). Você vai receber VÁRIOS ambientes de uma vez, cada um com suas constatações numeradas — algumas trazem entre colchetes a classificação já definida pela pessoa (ex: Avaria, Desgaste natural, Necessita manutenção, Sinal de umidade/infiltração, Estado de conservação, Funcionamento, Não testado, Não acessível). Responda SOMENTE com um JSON válido, sem markdown, sem texto antes ou depois, no formato exato: {"ambientes": [{"paragrafo": "...", "legendas": ["...", "..."]}, ...]}. O array "ambientes" deve ter EXATAMENTE um item por ambiente informado, NA MESMA ORDEM em que foram apresentados. Em cada item: "paragrafo" é um único parágrafo fluido (não uma lista) sintetizando TODAS as constatações daquele ambiente, mencionando de forma natural a classificação de cada uma quando informada (ex: "trata-se de condição associada a desgaste natural", "constatação classificada como avaria"), SEM atribuir culpa/responsabilidade ao locatário de forma direta, e mencionando que os itens estão registrados nas imagens indicadas daquele ambiente, que integram o laudo. "legendas" deve ter EXATAMENTE um item por constatação daquele ambiente, na mesma ordem, cada uma uma legenda curta (até 8 palavras) descrevendo objetivamente aquele item.'
    : 'Você é um(a) engenheiro(a) civil redigindo as seções de ambientes de um laudo técnico de vistoria de imóvel, seguindo o estilo formal de normas ABNT (somente as da lista conferida informada abaixo) e do IBAPE Nacional. Você vai receber VÁRIOS ambientes de uma vez, cada um com seus defeitos numerados. Responda SOMENTE com um JSON válido, sem markdown, sem texto antes ou depois, no formato exato: {"ambientes": [{"paragrafo": "...", "legendas": ["...", "..."]}, ...]}. O array "ambientes" deve ter EXATAMENTE um item por ambiente informado, NA MESMA ORDEM em que foram apresentados. Em cada item: "paragrafo" é um único parágrafo fluido (não uma lista, não repita os textos originais colados um atrás do outro) sintetizando TODOS os defeitos daquele ambiente, citando a norma ABNT mais pertinente quando fizer sentido, e mencionando que as não conformidades estão registradas nas imagens indicadas daquele ambiente, que integram o laudo. "legendas" deve ter EXATAMENTE um item por defeito daquele ambiente, na mesma ordem, cada um uma legenda curta (até 8 palavras) descrevendo objetivamente aquele defeito específico — se o texto original mencionar uma cor de adesivo/fita, a legenda deve citar essa cor (ex: "Adesivo verde indica cerâmica oca").';
  const mensagem = `${contextoExtra ? contextoExtra + '\n\n' : ''}Ambientes da vistoria:\n\n${blocos}`;
  const maxTokens = Math.min(8000, 500 + comDefeitos.length * 350);

  let ultimoErro = null;
  for (let tentativa = 1; tentativa <= 2; tentativa++) {
    let texto = '';
    try {
      texto = await chamarClaude({ model: 'claude-haiku-4-5-20251001', maxTokens, system: system + '\n\n' + REGRA_NORMAS, messages: [{ role: 'user', content: mensagem }] });
      const inicio = texto.indexOf('{');
      const fim = texto.lastIndexOf('}');
      if (inicio === -1 || fim === -1) throw new Error('Resposta não contém um JSON reconhecível.');
      const json = JSON.parse(texto.slice(inicio, fim + 1));
      if (!Array.isArray(json.ambientes) || json.ambientes.length !== comDefeitos.length) {
        throw new Error(`Quantidade de ambientes na resposta (${json.ambientes?.length}) não bate com o esperado (${comDefeitos.length}).`);
      }
      comDefeitos.forEach((a, idx) => {
        const item = json.ambientes[idx] || {};
        if (!item.paragrafo || typeof item.paragrafo !== 'string') throw new Error(`Ambiente "${a.ambiente}" sem parágrafo válido na resposta.`);
        if (!Array.isArray(item.legendas)) item.legendas = [];
        while (item.legendas.length < a.defeitos.length) item.legendas.push(null);
        item.legendas = item.legendas.slice(0, a.defeitos.length);
        resultado[a._indiceOriginal] = item;
      });
      return resultado;
    } catch (e) {
      ultimoErro = e;
      console.error(`Tentativa ${tentativa} de gerar resumo dos ambientes falhou:`, e.message, '| Resposta bruta da IA:', texto.slice(0, 800));
    }
  }
  throw ultimoErro;
}

// Tipos de documento que podem ser marcados como entregues no Laudo de Constatação.
// A ordem aqui é a ordem em que aparecem no laudo.
const DOCUMENTOS_PROJETOS = [
  'Projeto arquitetônico', 'Projeto estrutural', 'Projeto de fundações', 'Projeto hidrossanitário',
  'Projeto elétrico', 'Projeto de impermeabilização', 'Projeto de prevenção e combate a incêndio (PPCI)',
  'Projeto de águas pluviais', 'Projeto de gás', 'Projeto de climatização'
];
const DOCUMENTOS_GERAIS = [
  'Memorial descritivo', 'Manual de uso, operação e manutenção (manual do proprietário)',
  'Contrato', 'Matrícula do imóvel'
];

// Quebra um texto digitado em parágrafos (linha em branco ou quebra simples).
function emParagrafos(texto) {
  return String(texto || '').split(/\n+/).map(p => p.trim()).filter(Boolean);
}

// Monta o corpo do Laudo de Constatação, seguindo o modelo da G&G Engenharia:
// Identificação → Objetivo → Dados iniciais (endereço, localização, vistoria, data)
// → Planta → Documentos entregues → Anamnese → Metodologia → Registro fotográfico
// → (Outras observações) → Conclusão → Assinatura. A numeração das seções é
// calculada na hora, então seções opcionais (planta, observações) não deixam buraco.
function montarConteudoConstatacao({ dados, dataFmt, endFull, mapaParagraphs, plantaBase64, plantaMediaType, registrosParagraphs, obsGeral, conclusaoIA, executor, h }) {
  const { B, N, br, indent, secNum, subNum, bodyP, bodyBold, bulletP, defP, Paragraph, ImageRun, AlignmentType } = h;
  let n = 0;
  const sec = (titulo) => secNum(++n, titulo);
  const out = [];
  const nome = dados.nome || 'Cliente';

  // 1. IDENTIFICAÇÃO
  out.push(sec('IDENTIFICAÇÃO DO CONTRATANTE'));
  out.push(bodyP(`${nome}, portador(a) do CPF nº ${dados.cpf || '—'}${dados.telefone ? `, telefone ${dados.telefone}` : ''}, referente ao imóvel localizado na ${endFull || '—'}${dados.matricula ? `, matrícula nº ${dados.matricula}` : ''}, solicitou a elaboração do presente Laudo de Constatação, com a finalidade de registrar as condições aparentes do imóvel na data da inspeção, documentando visualmente as manifestações patológicas, danos e demais não conformidades observadas.`));
  out.push(br());

  // 2. OBJETIVO
  out.push(sec('OBJETIVO'));
  out.push(bodyP('O presente Laudo de Constatação tem por objetivo registrar, de forma técnica, objetiva e ilustrada, as condições aparentes do imóvel na data da inspeção, identificando manifestações patológicas, danos e demais inconformidades visíveis.'));
  out.push(bodyP('A avaliação restringe-se às condições perceptíveis por inspeção visual, não contemplando a realização de ensaios, testes de desempenho, medições instrumentais, abertura de elementos construtivos ou quaisquer procedimentos invasivos.'));
  out.push(bodyP('As constatações apresentadas neste documento representam o estado observado no momento da inspeção, servindo como registro técnico das evidências encontradas, sem estabelecer, de forma conclusiva, a origem, causa, responsabilidade ou extensão dos danos identificados, as quais dependem, quando necessário, de investigação complementar específica.'));
  out.push(br());

  // 3. DADOS INICIAIS
  const s3 = ++n;
  out.push(secNum(s3, 'DADOS INICIAIS'));
  out.push(subNum(`${s3}.1`, 'Identificação'));
  out.push(bodyBold('Endereço: ', endFull || '—'));
  if (dados.matricula) out.push(bodyBold('Matrícula do imóvel: ', dados.matricula));
  out.push(...mapaParagraphs);
  out.push(br());
  out.push(subNum(`${s3}.2`, 'Realização da vistoria'));
  if (textoRealizacaoVistoria(executor)) out.push(bodyP(textoRealizacaoVistoria(executor)));
  else { out.push(bodyP('Responsável: Engenheira Civil Gabriela Soares Borges')); out.push(bodyP('Registro CREA: 427760MG')); }
  out.push(br());
  out.push(subNum(`${s3}.3`, 'Data da vistoria'));
  out.push(bodyP(`A vistoria foi realizada dia ${dataFmt}.`));
  out.push(br());

  // 4. PLANTA (se houver)
  if (plantaBase64) {
    try {
      const plantaBuf = Buffer.from(plantaBase64, 'base64');
      out.push(sec('PLANTA DO IMÓVEL'));
      out.push(bodyP('Abaixo, apresenta-se croqui/planta da unidade vistoriada:'));
      out.push(new Paragraph({
        children: [new ImageRun({ data: plantaBuf, transformation: { width: 420, height: 300 }, type: plantaMediaType === 'image/png' ? 'png' : 'jpg' })],
        alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }
      }));
      out.push(new Paragraph({ children: [N('Planta / croqui do imóvel', 20)], alignment: AlignmentType.CENTER, spacing: { after: 160 } }));
      out.push(br());
    } catch (e) { console.error('Planta error:', e.message); }
  }

  // 5. DOCUMENTOS ENTREGUES
  const docs = Array.isArray(dados.documentos) ? dados.documentos : [];
  const projetos = DOCUMENTOS_PROJETOS.filter(d => docs.includes(d));
  const gerais = DOCUMENTOS_GERAIS.filter(d => docs.includes(d));
  const outrosDocs = emParagrafos(dados.documentosOutros);
  const faltando = [
    ...(projetos.length === 0 ? ['projetos'] : []),
    ...DOCUMENTOS_GERAIS.filter(d => !docs.includes(d)).map(d => d.split(' (')[0].toLowerCase())
  ];
  out.push(sec('DOCUMENTOS ENTREGUES'));
  if (projetos.length + gerais.length + outrosDocs.length === 0) {
    out.push(bodyP('Não foram disponibilizados pelo(a) contratante projetos, memorial descritivo, manual de uso, operação e manutenção, contrato ou matrícula do imóvel. Dessa forma, a análise baseou-se exclusivamente na inspeção visual realizada e nas informações prestadas na anamnese, não sendo possível confrontar as condições observadas com as especificações de projeto.'));
  } else {
    out.push(bodyP('Para a elaboração do presente laudo, foram disponibilizados pelo(a) contratante os seguintes documentos, considerados como referência na análise técnica:'));
    if (projetos.length) {
      out.push(new Paragraph({ children: [B('Projetos:', 20)], indent, spacing: { before: 60, after: 40 } }));
      projetos.forEach(d => out.push(bulletP(d + ';')));
    }
    if (projetos.length && (gerais.length || outrosDocs.length)) {
      out.push(new Paragraph({ children: [B('Demais documentos:', 20)], indent, spacing: { before: 60, after: 40 } }));
    }
    gerais.forEach(d => out.push(bulletP(d + ';')));
    outrosDocs.forEach(d => out.push(bulletP(d.replace(/[;,.]$/, '') + ';')));
    if (faltando.length) {
      out.push(bodyP(`Não foram disponibilizados: ${faltando.join(', ')}. As constatações relativas a esses itens baseiam-se exclusivamente na inspeção visual e nas informações prestadas pelo(a) contratante.`));
    }
  }
  out.push(br());

  // 6. ANAMNESE
  out.push(sec('ANAMNESE'));
  const anamnese = emParagrafos(dados.anamnese);
  if (anamnese.length) anamnese.forEach(p => out.push(bodyP(p)));
  else out.push(bodyP('Não foram prestadas informações adicionais pelo(a) contratante quanto ao histórico do imóvel, intervenções anteriores ou surgimento das manifestações observadas.'));
  out.push(br());

  // 7. METODOLOGIA
  out.push(sec('ELABORAÇÃO DO RELATÓRIO'));
  out.push(bodyP('A elaboração do presente laudo de constatação foi realizada com base na identificação dos elementos construtivos aparentes, sua localização dentro do imóvel e as manifestações patológicas visíveis no momento da inspeção.'));
  out.push(bodyP('Durante a vistoria, foram observados pontos de não conformidade, anomalias, falhas e manifestações patológicas que podem comprometer o desempenho, a segurança e a durabilidade dos sistemas e materiais da edificação.'));
  out.push(bodyP('A inspeção foi feita com base nos princípios estabelecidos pela ABNT NBR 16747:2020 – Inspeção predial: diretrizes, conceitos, terminologia, requisitos e procedimento, na ABNT NBR 5674:2024 – Manutenção de edificações, na ABNT NBR 15575 – Edificações habitacionais: Desempenho, e também conforme os conceitos definidos pelo IBAPE Nacional.'));
  out.push(bodyP('Considerando que alguns termos utilizados neste documento podem não ser de conhecimento geral, seguem abaixo os principais conceitos utilizados ao longo do relatório:'));
  out.push(defP('Anomalia', 'Irregularidade que compromete o desempenho de um elemento ou sistema da edificação. Pode ter origem no projeto, execução, uso ou manutenção inadequada.'));
  out.push(defP('Manifestação Patológica', 'Sinais visíveis de degradação, como fissuras, manchas, destacamentos, entre outros.'));
  out.push(defP('Agente de Degradação', 'Fatores (naturais, físicos ou químicos) que contribuem para a deterioração dos elementos construtivos.'));
  out.push(defP('Falha', 'Perda da função de um componente, seja por uso indevido, má execução ou falta de manutenção.'));
  out.push(defP('Desempenho', 'Comportamento da edificação e seus sistemas durante o uso, frente às solicitações normais esperadas ao longo de sua vida útil.'));
  out.push(defP('Vida Útil (VU)', 'Período em que um sistema ou componente deve cumprir suas funções, conforme previsto em projeto e respeitada sua manutenção adequada.'));
  out.push(defP('Plano de Manutenção', 'Documento técnico que organiza as ações necessárias de manutenção preventiva e corretiva de uma edificação.'));
  out.push(br());

  // 8. REGISTRO FOTOGRÁFICO
  out.push(sec('REGISTRO FOTOGRÁFICO DAS CONSTATAÇÕES'));
  out.push(bodyP(`A seguir, são apresentados os registros fotográficos das manifestações patológicas e não conformidades identificadas durante a vistoria realizada em ${dataFmt}, acompanhados da respectiva descrição técnica.`));
  out.push(br());
  out.push(...registrosParagraphs);

  // OUTRAS OBSERVAÇÕES
  if (obsGeral) {
    out.push(sec('OUTRAS OBSERVAÇÕES'));
    emParagrafos(obsGeral).forEach(p => out.push(bodyP(p)));
    out.push(br());
  }

  // CONCLUSÃO
  const conclusaoFixa = [
    'O presente laudo teve por objetivo o registro técnico das condições aparentes do imóvel vistoriado, com a identificação das manifestações patológicas, danos e não conformidades visíveis na data da inspeção. A análise foi desenvolvida com base em inspeção exclusivamente visual e em critérios técnicos consagrados na engenharia civil, observando-se as diretrizes da ABNT NBR 16747:2020, da ABNT NBR 5674:2024 e os requisitos de desempenho, durabilidade e segurança da ABNT NBR 15575.',
    'As manifestações registradas ao longo deste documento indicam comprometimento do desempenho dos sistemas construtivos afetados, em desacordo com os requisitos mínimos de habitabilidade, estanqueidade, segurança e durabilidade estabelecidos pela ABNT NBR 15575. A manutenção das condições observadas tende a promover a progressão das anomalias, com agravamento dos danos e redução da vida útil dos elementos afetados.',
    'Recomenda-se que as anomalias constatadas sejam objeto de diagnóstico complementar, quando necessário, e de intervenções corretivas executadas por profissional habilitado, com acompanhamento técnico e observância às normas aplicáveis a cada sistema, de modo a eliminar as causas das manifestações e não apenas seus sintomas, evitando a recorrência das falhas e intervenções futuras de maior complexidade, risco e custo.'
  ];
  const conclusao = (conclusaoIA && conclusaoIA.trim()) ? conclusaoIA.trim().split(/\n\s*\n/).map(p => p.trim()).filter(Boolean) : conclusaoFixa;
  out.push(sec('CONCLUSÃO'));
  conclusao.forEach(p => out.push(bodyP(p)));
  out.push(br());

  // ASSINATURA
  const ehExecutora = !!textoRealizacaoVistoria(executor);
  out.push(sec(ehExecutora ? 'ASSINATURAS' : 'ASSINATURA DO RESPONSÁVEL TÉCNICO'));
  out.push(bodyP(ehExecutora
    ? `Este Laudo Técnico de Constatação foi elaborado com base na vistoria realizada ${executor.titulo ? 'pela Eng. ' : 'por '}${executor.nome}, sob responsabilidade técnica da Engenheira Civil Gabriela Soares Borges, CREA 427760MG.`
    : 'Este Laudo Técnico de Constatação foi elaborado pela Engenheira Civil Gabriela Soares Borges.'));
  out.push(...blocoAssinaturas(executor, { N, B, br, Paragraph, AlignmentType }));
  return out;
}

// Textos de autoria do laudo. Quando quem fez a vistoria é uma executora, o laudo
// diz quem executou e mantém a Gabriela como responsável técnica (ART), e sai
// com duas assinaturas.
function textoRealizacaoVistoria(executor) {
  if (!executor || executor.ehResponsavelTecnica) return null;
  return `A vistoria foi realizada ${executor.titulo ? 'pela Eng. ' : 'por '}${executor.nome}, sob responsabilidade técnica da Eng. Gabriela Borges, conforme registro CREA 427760MG.`;
}
function blocoAssinaturas(executor, h) {
  const { N, B, br, Paragraph, AlignmentType } = h;
  const linha = () => new Paragraph({ children: [N('_______________________________', 20)], alignment: AlignmentType.CENTER, spacing: { before: 400, after: 60 } });
  const texto = (t, negrito) => new Paragraph({ children: [negrito ? B(t, 20) : N(t, 20)], alignment: AlignmentType.CENTER, spacing: { after: 40 } });
  const gabriela = [linha(), texto('Gabriela Soares Borges', true), texto('Engenheira Civil · CREA: 427760MG'), texto('Responsável técnica')];
  if (!executor || executor.ehResponsavelTecnica) return [linha(), texto('Gabriela Soares Borges'), texto('Engenheira Civil · CREA: 427760MG')];
  const exec = [linha(), texto(executor.nome, true)];
  if (executor.crea) exec.push(texto(`Engenheira Civil · CREA: ${executor.crea}`));
  exec.push(texto('Responsável pela execução da vistoria'));
  return [...exec, br(), ...gabriela];
}

async function gerarDocx(payload, executor) {
  const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, Footer } = require('docx');
  const { dados, tipoVistoria, obsGeral, registros, plantaBase64, plantaMediaType, mapaBase64, mapaMediaType, conclusaoIA } = payload;
  const ehLocacao = tipoVistoria === 'Locação';
  const ehConstatacao = tipoVistoria === 'Constatação';
  const modalidadeLocacao = ehLocacao ? (MODALIDADES_LOCACAO[dados.locacaoModalidade] || null) : null;

  const dataFmt = dados.data
    ? new Date(dados.data + 'T12:00:00').toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' })
    : new Date().toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' });

  const usarCapaConstatacao = ehConstatacao && fs.existsSync(COVER_CONSTATACAO_IMG);
  const coverImgData = fs.readFileSync(usarCapaConstatacao ? COVER_CONSTATACAO_IMG : COVER_IMG);
  const coverImgType = usarCapaConstatacao ? 'jpg' : 'png';
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
  const byAmbiente = {};
  registros.forEach(r => { if (!byAmbiente[r.ambiente]) byAmbiente[r.ambiente] = []; byAmbiente[r.ambiente].push(r); });

  // 1ª passagem: só organiza os dados (agrupa defeitos, calcula intervalo de
  // imagens de cada ambiente) — ainda sem chamar a IA nem montar o documento.
  let imgCounter = 1; // numeração contínua por TODO o laudo (não reinicia a cada ambiente/defeito)
  const ambientesInfo = [];
  for (const [ambiente, items] of Object.entries(byAmbiente)) {
    const porDefeito = {};
    const classificacaoPorDefeito = {};
    const ordemDefeitos = [];
    for (const item of items) {
      if (!porDefeito[item.defeito]) { porDefeito[item.defeito] = []; ordemDefeitos.push(item.defeito); }
      porDefeito[item.defeito].push(...item.fotos);
      if (item.classificacao && !classificacaoPorDefeito[item.defeito]) classificacaoPorDefeito[item.defeito] = item.classificacao;
    }
    const totalFotosAmbiente = ordemDefeitos.reduce((soma, d) => soma + porDefeito[d].length, 0);
    const inicioAmbiente = imgCounter;
    const fimAmbiente = imgCounter + totalFotosAmbiente - 1;
    imgCounter += totalFotosAmbiente;
    ambientesInfo.push({ ambiente, porDefeito, classificacaoPorDefeito, ordemDefeitos, inicioImg: inicioAmbiente, fimImg: fimAmbiente, defeitos: ordemDefeitos, _indiceOriginal: ambientesInfo.length });
  }

  // Uma ÚNICA chamada de IA cobrindo TODOS os ambientes de uma vez (em vez de
  // uma por ambiente) — bem mais barato. A pedido da cliente, o laudo NUNCA deve
  // sair no formato antigo (parágrafos repetidos); se a IA falhar mesmo depois
  // de tentar 2 vezes, a geração do laudo INTEIRO é interrompida aqui, e o app
  // do celular já trata isso (fila de reenvio automático) — nunca perde nada.
  const resumosPorAmbiente = await gerarResumosDosAmbientes(ambientesInfo, tipoVistoria, modalidadeLocacao ? modalidadeLocacao.contextoIA : '');

  // 2ª passagem: agora sim monta o documento de verdade, usando os parágrafos e
  // legendas já prontos, na mesma ordem/numeração calculada na 1ª passagem.
  for (const info of ambientesInfo) {
    const { ambiente, porDefeito, ordemDefeitos } = info;
    const resumoIA = resumosPorAmbiente[info._indiceOriginal];
    let imgAtual = info.inicioImg;

    registrosParagraphs.push(new Paragraph({
      children: [B(ambiente, 20)], indent, spacing: { before: 200, after: 100 },
      keepNext: true // não deixa o nome do ambiente sozinho no pé da página
    }));

    // Parágrafo descritivo por ambiente: só no Laudo de Constatação. Na Vistoria,
    // Revistoria e Locação cada ambiente traz só as fotos com legenda, e a análise
    // técnica fica concentrada na Conclusão.
    if (ehConstatacao) {
      registrosParagraphs.push(new Paragraph({
        children: [N(resumoIA.paragrafo, 20)], indent, spacing: { after: 140 }, alignment: AlignmentType.JUSTIFIED, keepNext: true
      }));
    }

    let idxDefeito = 0;
    for (const defeito of ordemDefeitos) {
      const fotosDoGrupo = porDefeito[defeito];
      const legendaCurta = resumoIA.legendas[idxDefeito];

      for (let i = 0; i < fotosDoGrupo.length; i++) {
        const foto = fotosDoGrupo[i];
        const rotulo = legendaCurta ? `Imagem ${imgAtual} - ${legendaCurta}` : `Imagem ${imgAtual}`;
        registrosParagraphs.push(new Paragraph({
          children: [N(rotulo, 20)], alignment: AlignmentType.CENTER, spacing: { after: 60 }, keepNext: true
        }));
        try {
          const imgBuf = Buffer.from(foto.base64, 'base64');
          const imgType = foto.mediaType === 'image/png' ? 'png' : 'jpg';
          registrosParagraphs.push(new Paragraph({
            children: [new ImageRun({ data: imgBuf, transformation: { width: 400, height: 280 }, type: imgType })],
            alignment: AlignmentType.CENTER, spacing: { before: 80, after: 160 }
          }));
        } catch(e) { console.error('Img error:', e.message); }
        imgAtual++;
      }
      idxDefeito++;
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
  const elaboracaoTexto1 = ehLocacao
    ? 'A elaboração do presente relatório de vistoria de locação foi realizada com base na identificação dos elementos construtivos e do mobiliário aparentes, sua localização dentro do imóvel e as condições de conservação e funcionamento observadas no momento da inspeção.'
    : 'A elaboração do presente relatório de vistoria técnica de recebimento da unidade habitacional foi realizada com base na identificação dos elementos construtivos aparentes, sua localização dentro do imóvel e as manifestações patológicas visíveis no momento da inspeção.';
  const elaboracaoTexto2 = ehLocacao
    ? 'Durante a vistoria, foram observadas as condições de conservação, funcionamento e acabamento dos ambientes, sendo cada constatação classificada quanto à sua natureza (avaria, desgaste natural, necessidade de manutenção, entre outras), conforme especificado em cada registro.'
    : 'Durante a vistoria, foram observados diversos pontos de não conformidade, falhas de acabamento, anomalias e possíveis vícios construtivos que podem comprometer o desempenho esperado dos sistemas e materiais.';
  const elaboracaoTexto3 = 'A inspeção foi feita com base nos princípios estabelecidos pela ABNT NBR 16747:2020 – Inspeção predial: diretrizes, conceitos, terminologia, requisitos e procedimento, na ABNT NBR 5674:2024 – Manutenção de edificações, e também conforme os conceitos definidos pelo IBAPE Nacional' + (ehLocacao ? ', sem prejuízo das disposições da Lei nº 8.245/1991 (Lei do Inquilinato).' : '.');
  const elaboracaoTexto4 = 'Considerando que alguns termos utilizados neste documento podem não ser de conhecimento geral, seguem abaixo os principais conceitos utilizados ao longo do relatório:';
  const elaboracaoTexto5 = ehLocacao
    ? 'A presente vistoria considera, sem prejuízo das demais normas técnicas aplicáveis, as disposições da Lei nº 8.245/1991 (Lei do Inquilinato), que rege a relação entre locador e locatário quanto ao estado de entrega e restituição do imóvel.'
    : 'A unidade inspecionada apresenta diversas não conformidades visuais. Tais ocorrências indicam ausência de cuidados na execução final e comprometem o recebimento do imóvel em condições ideais de entrega.';
  const elaboracaoTexto6 = ehLocacao ? null
    : 'A recomendação técnica é que todas as anomalias listadas neste relatório sejam corrigidas antes da conclusão da entrega da unidade ao proprietário, garantindo o desempenho mínimo esperado e evitando prejuízos futuros.';

  // Parágrafo de proteção legal — SEMPRE incluído (literal, não passa pela IA)
  // na Conclusão de laudos de Locação, conforme aprovado.
  const TEXTO_LEI_LOCACAO = 'Nos termos da Lei nº 8.245/1991 (Lei do Inquilinato), cabe ao locador entregar o imóvel em estado de servir ao uso a que se destina, respondendo pelos vícios ou defeitos anteriores à locação, enquanto cabe ao locatário restituí-lo, ao final do contrato, no estado em que o recebeu, ressalvadas as deteriorações decorrentes do uso normal.';
  const TEXTO_PROTECAO_LOCACAO = 'A presente vistoria tem por finalidade registrar o estado aparente de conservação, funcionamento e características dos elementos acessíveis do imóvel na data da inspeção, servindo como documento de caracterização para comparação entre as condições observadas no início e no término da ocupação. As constatações apresentadas não constituem, isoladamente, atribuição de responsabilidade pela origem das manifestações, devendo eventual responsabilização considerar a natureza do dano, seu histórico, condições preexistentes, desgaste natural, uso, manutenção e demais evidências disponíveis.';

  // Se o front-end mandou uma conclusão gerada por IA (baseada nos defeitos reais), usa ela.
  // Caso contrário (IA falhou, ou não foi chamada), cai no texto fixo genérico de sempre.
  const conclusaoTextoFixo = ehLocacao ? [
    'Com base na vistoria técnica realizada no imóvel objeto da relação locatícia, foram registradas as condições de conservação, funcionamento e acabamento dos ambientes e elementos vistoriados, conforme descrito e documentado ao longo deste relatório técnico.'
  ] : [
    'Com base na vistoria técnica realizada na unidade habitacional, constatou-se a presença de não conformidades construtivas, falhas de acabamento e inconformidades funcionais distribuídas nos ambientes inspecionados, conforme descrito e documentado ao longo deste relatório técnico. As manifestações observadas incluem irregularidades em revestimentos, falhas de rejuntamento, defeitos em pintura, problemas em esquadrias, portas, elementos hidráulicos, acabamentos e demais sistemas construtivos aparentes.',
    'Os defeitos identificados evidenciam deficiência nos processos executivos e no controle de qualidade durante as etapas de acabamento e entrega da unidade, não sendo compatíveis com o padrão esperado para um imóvel novo. Ainda que parte das inconformidades apresente caráter predominantemente estético, diversas manifestações podem comprometer a durabilidade dos materiais, o desempenho dos sistemas construtivos, a estanqueidade, a funcionalidade dos ambientes e a vida útil da edificação ao longo do tempo.',
    'Conforme os princípios estabelecidos pela ABNT NBR 15575, a edificação deve atender aos requisitos mínimos de desempenho relacionados à segurança, habitabilidade, funcionalidade e durabilidade. Da mesma forma, os serviços executivos e acabamentos devem seguir padrões adequados de qualidade e conformidade técnica, observando as boas práticas construtivas e as normas aplicáveis a cada sistema construtivo. As anomalias constatadas neste relatório demonstram inconformidades em relação a tais requisitos, tornando tecnicamente recomendável a correção integral dos itens apontados.',
    'Dessa forma, conclui-se que todas as não conformidades registradas neste documento devem ser devidamente corrigidas pela construtora/responsável técnico antes da aceitação definitiva do imóvel, garantindo o adequado desempenho dos sistemas, a preservação da vida útil dos materiais e o padrão de qualidade esperado para a edificação. Recomenda-se ainda que os reparos sejam executados com acompanhamento técnico e observância aos procedimentos normativos aplicáveis, a fim de evitar recorrência das falhas identificadas.'
  ];
  const conclusaoTexto = (conclusaoIA && conclusaoIA.trim())
    ? conclusaoIA.trim().split(/\n\s*\n/).map(p => p.trim()).filter(Boolean)
    : conclusaoTextoFixo;

  // ─── LAUDO DE CONSTATAÇÃO (estrutura própria) ───────────
  const childrenConstatacao = ehConstatacao ? montarConteudoConstatacao({
    dados, dataFmt, endFull, mapaParagraphs, plantaBase64, plantaMediaType, registrosParagraphs, obsGeral, conclusaoIA, executor,
    h: { B, N, br, indent, secNum, subNum, bodyP, bodyBold, bulletP, defP, Paragraph, ImageRun, AlignmentType }
  }) : null;

  const doc = new Document({
    sections: [
      // CAPA
      {
        properties: { page: { size: { width: 11910, height: 16840 }, margin: { top: 0, right: 0, bottom: 0, left: 0 } } },
        children: [
          new Paragraph({
            children: [new ImageRun({ data: coverImgData, transformation: { width: 794, height: 1123 }, type: coverImgType })],
            alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }
          })
        ]
      },
      // CONTEÚDO
      {
        properties: { page: { size: { width: 11910, height: 16840 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
        footers: { default: makeFooter() },
        children: childrenConstatacao || [
          // 1. IDENTIFICAÇÃO
          secNum(1, 'IDENTIFICAÇÃO DO CONTRATANTE'),
          bodyP(ehLocacao
            ? `${dados.nome || 'Cliente'}, portador(a) do CPF nº ${dados.cpf || '—'}${dados.telefone ? `, telefone ${dados.telefone}` : ''}, solicitou a elaboração do presente relatório de ${modalidadeLocacao ? modalidadeLocacao.rotulo : 'vistoria de locação de imóvel'}, com o objetivo de registrar as condições da unidade na data da vistoria, para fins de caracterização do estado de conservação, funcionamento e acabamentos, servindo como documento de referência para a relação locatícia.`
            : `${dados.nome || 'Cliente'}, portador(a) do CPF nº ${dados.cpf || '—'}${dados.telefone ? `, telefone ${dados.telefone}` : ''}, solicitou a elaboração do presente relatório de vistoria de imóvel, com o objetivo de registrar as condições da unidade no momento da entrega, identificando eventuais inconformidades aparentes e falhas de execução visíveis.`),
          br(),

          // 2. OBJETIVO
          secNum(2, 'OBJETIVO'),
          ...(modalidadeLocacao ? [
            bodyBold('Modalidade da vistoria: ', dados.locacaoModalidade),
            bodyP(modalidadeLocacao.objetivo),
            bodyP('A avaliação foi conduzida de forma objetiva e detalhada, com base nas diretrizes técnicas aplicáveis, sem prejuízo das disposições da Lei nº 8.245/1991 (Lei do Inquilinato), que rege as relações locatícias.')
          ] : [bodyP(ehLocacao
            ? `Este relatório tem como finalidade documentar, de forma objetiva e detalhada, as condições do imóvel na data da vistoria de locação, identificando o estado de conservação, funcionamento e características dos elementos acessíveis do imóvel, servindo como documento de caracterização para eventual comparação entre as condições observadas no início e no término da ocupação. A avaliação foi conduzida com base nas diretrizes técnicas aplicáveis, sem prejuízo das disposições da Lei nº 8.245/1991 (Lei do Inquilinato), que rege as relações locatícias.`
            : (tipoVistoria === 'Revistoria'
                ? `Este relatório tem como finalidade documentar, de forma objetiva e detalhada, as condições do imóvel na data da revistoria, verificando os itens anteriormente apontados em vistoria prévia e registrando as não conformidades que permanecem pendentes de correção. A avaliação foi conduzida com base nas diretrizes estabelecidas pelas normas técnicas vigentes, incluindo as NBR (Normas Brasileiras) aplicáveis e os referenciais do PBQP-H (Programa Brasileiro da Qualidade e Produtividade no Habitat), assegurando a análise dos padrões de qualidade, segurança, funcionalidade e durabilidade do empreendimento.`
                : `Este relatório tem como finalidade documentar, de forma objetiva e detalhada, as condições do imóvel ${tipoVistoria === 'Imóvel Novo' ? 'novo ' : ''}na data da vistoria, identificando eventuais não conformidades em acabamentos, instalações elétricas e hidráulicas, estrutura e funcionalidade dos ambientes. A avaliação foi conduzida seguindo as diretrizes estabelecidas pelas normas técnicas vigentes, incluindo as NBR (Normas Brasileiras) e os referenciais do PBQP-H (Programa Brasileiro da Qualidade e Produtividade no Habitat), assegurando que os padrões de qualidade, segurança, funcionalidade e durabilidade do empreendimento sejam observados.`))]),
          br(),

          // 3. DADOS INICIAIS
          secNum(3, 'DADOS INICIAIS'),
          subNum('3.1', 'Identificação'),
          bodyBold('Empreendimento: ', dados.empreendimento || '—'),
          bodyBold('Endereço: ', endFull),
          ...mapaParagraphs,
          br(),
          subNum('3.2', 'Realização da vistoria'),
          ...(textoRealizacaoVistoria(executor)
            ? [bodyP(textoRealizacaoVistoria(executor))]
            : [bodyP('Responsável: Engenheira Civil Gabriela Soares Borges'), bodyP('Registro CREA: 427760MG')]),
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
          bodyP(ehLocacao
            ? 'Durante a vistoria, foram inspecionados os acabamentos, instalações elétricas e hidráulicas, mobiliário (quando aplicável), funcionalidade dos ambientes e demais itens que compõem o imóvel, registrando-se o estado de conservação e eventuais não conformidades para fins de caracterização da unidade na presente relação locatícia.'
            : 'Durante a vistoria, foram inspecionados os acabamentos, instalações elétricas e hidráulicas, funcionalidade dos ambientes e demais itens que compõem o imóvel, registrando-se eventuais não conformidades para que sejam corrigidas conforme os padrões de qualidade estabelecidos pela construtora.'),
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
          ...(ehLocacao ? [
            defP('Anomalia', 'Irregularidade que compromete o desempenho de um elemento ou sistema da edificação. Pode ter origem no projeto, execução, uso ou manutenção inadequada.'),
            defP('Manifestação Patológica', 'Sinais visíveis de degradação, como fissuras, manchas, destacamentos, entre outros.'),
            defP('Desgaste Natural', 'Deterioração decorrente do uso normal do imóvel ao longo do tempo, não configurando dano imputável ao ocupante.'),
            defP('Desempenho', 'Comportamento da edificação e seus sistemas durante o uso, frente às solicitações normais esperadas ao longo de sua vida útil.'),
            defP('Vida Útil (VU)', 'Período em que um sistema ou componente deve cumprir suas funções, conforme previsto em projeto e respeitada sua manutenção adequada.'),
          ] : [
            defP('Anomalia', 'Irregularidade que compromete o desempenho de um elemento ou sistema da edificação. Pode ter origem no projeto, execução, uso ou manutenção inadequada.'),
            defP('Manifestação Patológica', 'Sinais visíveis de degradação, como fissuras, manchas, destacamentos, entre outros.'),
            defP('Agente de Degradação', 'Fatores (naturais, físicos ou químicos) que contribuem para a deterioração dos elementos construtivos.'),
            defP('Falha', 'Perda da função de um componente, seja por uso indevido, má execução ou falta de manutenção.'),
            defP('Desempenho', 'Comportamento da edificação e seus sistemas durante o uso, frente às solicitações normais esperadas ao longo de sua vida útil.'),
            defP('Vida Útil (VU)', 'Período em que um sistema ou componente deve cumprir suas funções, conforme previsto em projeto e respeitada sua manutenção adequada.'),
            defP('Plano de Manutenção', 'Documento técnico que organiza as ações necessárias de manutenção preventiva e corretiva de uma edificação.'),
          ]),
          br(),
          bodyP(elaboracaoTexto5),
          br(),
          ...(elaboracaoTexto6 ? [bodyP(elaboracaoTexto6), br()] : []),

          // 7. REGISTROS
          secNum(6 + numOffset, ehLocacao ? 'REGISTRO DE CONSTATAÇÕES DA VISTORIA' : 'REGISTRO DE NÃO CONFORMIDADES DA VISTORIA'),
          bodyP(ehLocacao
            ? `A seguir, são apresentados, por ambiente, os registros fotográficos das constatações identificadas durante a vistoria no dia ${dataFmt}, cada um acompanhado de legenda identificando o item registrado. A análise técnica consolidada dessas constatações é apresentada na Conclusão deste relatório.`
            : `A seguir, são apresentados, por ambiente, os registros fotográficos das não conformidades identificadas durante a vistoria no dia ${dataFmt}, cada um acompanhado de legenda identificando a não conformidade registrada. A análise técnica consolidada dessas não conformidades, com a fundamentação normativa aplicável, é apresentada na Conclusão deste relatório.`),
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
          ...(ehLocacao ? [bodyP(TEXTO_LEI_LOCACAO), bodyP(modalidadeLocacao && modalidadeLocacao.semComparacao
            ? TEXTO_PROTECAO_LOCACAO.replace(', servindo como documento de caracterização para comparação entre as condições observadas no início e no término da ocupação', ', servindo como documento de caracterização do estado do imóvel naquele momento')
            : TEXTO_PROTECAO_LOCACAO)] : []),
          br(),

          // ASSINATURA
          secNum(obsGeral ? 9 + numOffset : 8 + numOffset, textoRealizacaoVistoria(executor) ? 'ASSINATURAS' : 'ASSINATURA DO RESPONSÁVEL'),
          ...blocoAssinaturas(executor, { N, B, br, Paragraph, AlignmentType }),

        ]
      }
    ]
  });

  return Packer.toBuffer(doc);
}

// Conclusão ROBUSTA do Laudo de Constatação: usa os defeitos de cada ambiente,
// a anamnese e os documentos entregues pra redigir uma conclusão no estilo
// pericial, organizada por sistema construtivo e fundamentada nas normas ABNT.
async function gerarConclusaoConstatacao({ registros, obsGeral, dados }) {
  const porAmbiente = {};
  (registros || []).forEach(r => { (porAmbiente[r.ambiente] = porAmbiente[r.ambiente] || []).push(r.defeito); });
  const listaDefeitos = Object.entries(porAmbiente)
    .map(([amb, defs]) => `${amb}:\n${[...new Set(defs)].map(d => `  - ${d}`).join('\n')}`)
    .join('\n') || 'Nenhuma constatação registrada.';
  const docs = Array.isArray(dados.documentos) ? dados.documentos : [];
  const docsTexto = [...docs, ...emParagrafos(dados.documentosOutros)].join('; ') || 'Nenhum documento foi entregue.';

  const system = `Você é um(a) engenheiro(a) civil perito(a), redigindo a seção CONCLUSÃO de um LAUDO DE CONSTATAÇÃO de imóvel, elaborado por inspeção exclusivamente visual (sem ensaios, cálculos ou procedimentos invasivos). A conclusão deve ser ROBUSTA, técnica e fundamentada em normas, no estilo de laudo pericial brasileiro.

Estrutura obrigatória (em parágrafos corridos, sem títulos):
1. Parágrafo de abertura: objetivo do laudo, natureza visual da inspeção e as normas que fundamentam a análise (escolha as pertinentes às constatações).
2. Um parágrafo para CADA grupo de manifestações/sistema construtivo afetado (ex: umidade e infiltrações; cobertura e telhado; estrutura de concreto e armaduras; fissuras em alvenaria/lajes; revestimentos e pinturas; impermeabilização; instalações hidrossanitárias; instalações elétricas; esquadrias; intervenções/reparos executados). Em cada um: descreva o que foi constatado, explique tecnicamente o mecanismo provável e as consequências (usando linguagem cautelosa: "compatível com", "indícios de", "sugere"), e cite a norma ABNT aplicável e o requisito que deixa de ser atendido.
3. Quando a anamnese relatar intervenções anteriores, relacione-as tecnicamente com o constatado (ex: reparo paliativo que não atua na causa), sem afirmar culpa de forma categórica.
4. Se faltarem projetos ou documentos relevantes, registre a limitação que isso impõe à análise.
5. Parágrafo de riscos (segurança dos usuários, progressão das patologias, redução da vida útil) quando pertinente.
6. Parágrafo final conclusivo: as condições observadas não atendem aos requisitos de desempenho, segurança, estanqueidade e/ou durabilidade aplicáveis; recomendação de diagnóstico complementar/intervenções corretivas por profissional habilitado, atuando na causa raiz.

Normas de referência (cite só as pertinentes, com número correto): ${NORMAS_REFERENCIA}

${REGRA_NORMAS}

Regras: NÃO invente constatações, medidas, datas ou fatos que não estejam nas informações recebidas. Português formal, impessoal. Entre 6 e 10 parágrafos. Retorne SOMENTE os parágrafos, separados por uma linha em branco, sem títulos, sem markdown, sem numeração, sem listas.`;

  const mensagem = `Contratante: ${dados.nome || 'não informado'}
Imóvel: ${[dados.endereco, dados.cidade].filter(Boolean).join(', ') || 'não informado'}

ANAMNESE (relato do contratante):
${(dados.anamnese || '').trim() || 'Não informada.'}

DOCUMENTOS ENTREGUES:
${docsTexto}

CONSTATAÇÕES REGISTRADAS NA VISTORIA (por ambiente):
${listaDefeitos}

OUTRAS OBSERVAÇÕES:
${(obsGeral || '').trim() || 'Nenhuma'}

Redija a conclusão do laudo.`;

  return chamarClaudeTecnico({ maxTokens: 4000, system, messages: [{ role: 'user', content: mensagem }] });
}

// Conclusão ROBUSTA dos laudos de Vistoria, Revistoria e Locação. Como esses
// laudos não trazem mais parágrafo por ambiente (só fotos com legenda), toda a
// análise técnica fica aqui: organizada por sistema construtivo, com o mecanismo
// provável de cada falha, consequências e a norma aplicável.

async function gerarConclusaoVistoria({ registros, obsGeral, tipoVistoria, dados }) {
  const ehLocacao = tipoVistoria === 'Locação';
  const ehRevistoria = tipoVistoria === 'Revistoria';
  const modalidade = ehLocacao ? MODALIDADES_LOCACAO[dados.locacaoModalidade] : null;

  const porAmbiente = {};
  (registros || []).forEach(r => {
    const linha = r.classificacao ? `${r.defeito} [classificação: ${r.classificacao}]` : r.defeito;
    (porAmbiente[r.ambiente] = porAmbiente[r.ambiente] || []).push(linha);
  });
  const listaDefeitos = Object.entries(porAmbiente)
    .map(([amb, defs]) => `${amb}:\n${[...new Set(defs)].map(d => `  - ${d}`).join('\n')}`)
    .join('\n') || 'Nenhuma constatação registrada.';

  const estrutura = `Estrutura (em parágrafos corridos, sem títulos):
1. Abertura: o que foi vistoriado, a natureza visual da inspeção e as normas que fundamentam a análise (escolha as pertinentes).
2. Um parágrafo para CADA sistema/grupo afetado, agrupando os itens de todos os ambientes (ex: revestimentos de piso e parede e rejuntes; pintura e acabamentos; paredes, fissuras, prumo e esquadro; esquadrias, portas e janelas; instalações hidrossanitárias, louças, metais e ralos; instalações elétricas; caimentos e estanqueidade/impermeabilização; cobertura, telhado e alçapão; garagem e fachada; limpeza). Em cada um: descreva o conjunto do que foi registrado (cite os ambientes quando ajudar), explique tecnicamente as causas prováveis e as consequências (linguagem cautelosa: "compatível com", "indícios de"), e cite a norma ABNT aplicável e o requisito que deixa de ser atendido. Dê mais peso ao que afeta segurança, estanqueidade, funcionamento e durabilidade do que ao que é só estético.
3. Um parágrafo de priorização: o que exige correção mais urgente (segurança, risco elétrico, infiltrações, estrutura) e o que é de acabamento.`;

  let papel, finalizacao;
  if (ehLocacao) {
    papel = `Você é um(a) engenheiro(a) civil redigindo a seção CONCLUSÃO de um laudo técnico de VISTORIA DE LOCAÇÃO de imóvel${dados.locacaoModalidade ? ` — modalidade: ${dados.locacaoModalidade.toUpperCase()}` : ''}. ${modalidade ? modalidade.contextoIA : ''} As constatações podem vir com classificação (avaria, desgaste natural, necessita manutenção, sinal de umidade, estado de conservação, funcionamento, não testado, não acessível) — use-as ao agrupar e comentar. NUNCA atribua culpa ou responsabilidade ao locatário ou ao locador de forma categórica; trate como registro técnico do estado do imóvel. NÃO mencione construtora nem aceitação de obra. NÃO inclua menção à Lei 8.245/1991 nem parágrafo de proteção legal — isso é acrescentado automaticamente depois do seu texto.`;
    finalizacao = '4. Parágrafo final: síntese do estado geral de conservação e funcionamento do imóvel na data da vistoria e recomendação de manutenção dos itens apontados.';
  } else if (ehRevistoria) {
    papel = 'Você é um(a) engenheiro(a) civil redigindo a seção CONCLUSÃO de um laudo técnico de REVISTORIA de imóvel (nova inspeção para verificar se as não conformidades apontadas na vistoria anterior foram corrigidas). As constatações listadas são as não conformidades que PERMANECEM pendentes ou foram identificadas nesta revistoria.';
    finalizacao = '4. Parágrafo final: as não conformidades listadas permanecem pendentes de correção, não sendo recomendável a aceitação definitiva do imóvel até sua correção integral pela construtora/responsável técnico, com acompanhamento técnico e nova verificação após os reparos.';
  } else {
    papel = 'Você é um(a) engenheiro(a) civil redigindo a seção CONCLUSÃO de um laudo técnico de VISTORIA DE RECEBIMENTO de imóvel (entrega de unidade pela construtora), seguindo a ABNT NBR 16747:2020, a ABNT NBR 15575 e os conceitos do IBAPE Nacional.';
    finalizacao = '4. Parágrafo final: conclusão de que as não conformidades registradas devem ser corrigidas pela construtora/responsável técnico antes da aceitação definitiva do imóvel, com acompanhamento técnico e observância às normas aplicáveis; quando pertinente, mencione que eventuais vícios ocultos que se manifestem depois seguem os prazos de garantia da ABNT NBR 17170:2022, e que vícios aparentes devem ser formalmente comunicados à construtora.';
  }

  const system = `${papel}

A conclusão deve ser ROBUSTA, técnica e fundamentada em normas — ela é a única parte do laudo com análise escrita (os ambientes trazem apenas fotos com legenda curta).

${estrutura}
${finalizacao}

Normas de referência (cite só as pertinentes, com número correto): ${NORMAS_REFERENCIA}

${REGRA_NORMAS}

Regras: NÃO invente constatações, medidas, datas, marcas ou fatos que não estejam nas informações recebidas. Português formal e impessoal. Entre 5 e 9 parágrafos, proporcional à quantidade de constatações. Retorne SOMENTE os parágrafos, separados por uma linha em branco, sem títulos, sem markdown, sem numeração, sem listas.`;

  const mensagem = `Tipo de laudo: ${tipoVistoria || 'Vistoria'}${dados.locacaoModalidade ? ` (${dados.locacaoModalidade})` : ''}
Imóvel: ${[dados.empreendimento, dados.endereco, dados.bloco, dados.apto, dados.cidade].filter(Boolean).join(', ') || 'não informado'}
${dados.construtora ? `Construtora: ${dados.construtora}\n` : ''}
CONSTATAÇÕES REGISTRADAS NA VISTORIA (por ambiente):
${listaDefeitos}

OUTRAS OBSERVAÇÕES:
${(obsGeral || '').trim() || 'Nenhuma'}

Redija a conclusão do laudo.`;

  return chamarClaudeTecnico({ maxTokens: 4000, system, messages: [{ role: 'user', content: mensagem }] });
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

// ══════════════════════════════════════════════════════════════
// CONTAS DE USUÁRIO
// • ADM (Gabriela): entra com o usuário "Gabriela" e a senha de APP_PASSWORD
//   (a mesma de sempre). Vê os próprios dados e, pela guia "Ver conta", os
//   dados de qualquer executora.
// • Executoras: cadastradas pela ADM na tela /admin/usuarios. Cada uma vê só
//   os próprios laudos, agenda, vistorias em andamento e financeiro.
// As senhas ficam só no Volume do Railway, criptografadas (scrypt) — nunca no
// código (o repositório é público).
// ══════════════════════════════════════════════════════════════
const crypto = require('crypto');
const USUARIOS_FILE = path.join(OUTPUT_DIR, 'usuarios.json');
const SEGREDO_FILE = path.join(OUTPUT_DIR, '.segredo_sessao');
const ADMIN_ID = 'gabriela';
const ADMIN = {
  id: ADMIN_ID, usuario: 'Gabriela', nome: 'Gabriela Soares Borges', nomeCurto: 'Gabriela Borges',
  titulo: 'Eng.', crea: '427760MG', papel: 'admin'
};
const TERMO_VERSAO = '2026-10';

function segredoSessao() {
  try { return fs.readFileSync(SEGREDO_FILE, 'utf8').trim(); } catch {}
  const novo = crypto.randomBytes(32).toString('hex');
  fs.writeFileSync(SEGREDO_FILE, novo);
  return novo;
}
const SEGREDO = segredoSessao();

function lerUsuarios() { try { return JSON.parse(fs.readFileSync(USUARIOS_FILE, 'utf8')); } catch { return []; } }
function salvarUsuarios(lista) {
  const tmp = USUARIOS_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(lista, null, 2));
  fs.renameSync(tmp, USUARIOS_FILE);
}
// "Victória", "victoria", " VICTORIA " → "victoria"
const normalizarUsuario = (u) => String(u || '').normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();
function hashSenha(senha, salt) { return crypto.scryptSync(String(senha), salt, 64).toString('hex'); }
function conferirSenha(senha, salt, hash) {
  const a = Buffer.from(hashSenha(senha, salt), 'hex'), b = Buffer.from(hash, 'hex');
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}
function senhaForte(s) { return typeof s === 'string' && s.length >= 8 && /[A-Za-z]/.test(s) && /\d/.test(s); }

// Sessão: cookie assinado com "id.versaoDaSenha.validade" — trocar/redefinir a
// senha derruba as sessões antigas daquela pessoa.
function criarToken(id, versao) {
  const corpo = `${id}.${versao || 0}.${Date.now() + 365 * 24 * 3600 * 1000}`;
  const assinatura = crypto.createHmac('sha256', SEGREDO).update(corpo).digest('hex');
  return `${corpo}.${assinatura}`;
}
function lerToken(token) {
  const partes = String(token || '').split('.');
  if (partes.length !== 4) return null;
  const corpo = partes.slice(0, 3).join('.');
  const esperado = crypto.createHmac('sha256', SEGREDO).update(corpo).digest('hex');
  if (esperado.length !== partes[3].length || !crypto.timingSafeEqual(Buffer.from(esperado), Buffer.from(partes[3]))) return null;
  if (Number(partes[2]) < Date.now()) return null;
  return { id: partes[0], versao: Number(partes[1]) };
}
function cookieSessao(id, versao) {
  return `engcheck_sessao=${criarToken(id, versao)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=31536000`;
}

// Quem está usando o sistema nesta requisição (ou null se ninguém logado)
function usuarioDaSessao(req) {
  if (!APP_PASSWORD) return { ...ADMIN }; // sem senha configurada: comportamento antigo, sem login
  const cookies = parseCookies(req.headers.cookie);
  const t = lerToken(cookies['engcheck_sessao']);
  if (t) {
    if (t.id === ADMIN_ID) return { ...ADMIN };
    const u = lerUsuarios().find(x => x.id === t.id);
    if (u && u.ativo !== false && (u.senhaVersao || 0) === t.versao) return { ...u, papel: 'executora' };
    return null;
  }
  // compatibilidade: sessões antigas da ADM (antes das contas) continuam valendo
  if (cookies['engcheck_auth'] && cookies['engcheck_auth'] === APP_PASSWORD) return { ...ADMIN, sessaoAntiga: true };
  return null;
}
function pendenciaDoUsuario(u) {
  if (!u || u.papel === 'admin') return null;
  if (!u.termo || u.termo.versao !== TERMO_VERSAO) return 'termo';
  if (u.trocarSenha) return 'senha';
  return null;
}

// Qual conta a ADM está visualizando (guia "Ver conta"): o id de uma executora,
// 'todas', ou a própria (padrão). Executoras sempre veem só a própria.
function contaVisualizada(req, u) {
  if (!u || u.papel !== 'admin') return u ? u.id : null;
  const ver = parseCookies(req.headers.cookie)['engcheck_ver'];
  if (ver === 'todas') return 'todas';
  if (ver && ver !== ADMIN_ID && lerUsuarios().some(x => x.id === ver)) return ver;
  return ADMIN_ID;
}
const donoDoItem = (item) => (item && item.contaId) || ADMIN_ID; // itens antigos = da ADM
// Filtro das listas (histórico, agenda, rascunhos) conforme quem está logado e a conta visualizada
function filtroDeConta(req, u) {
  const ver = contaVisualizada(req, u);
  if (ver === 'todas') return () => true;
  return (item) => donoDoItem(item) === ver;
}
// Pode mexer neste item? ADM pode tudo; executora só no que é dela
const podeMexer = (u, item) => !!u && (u.papel === 'admin' || donoDoItem(item) === u.id);

function nomeDaConta(id) {
  if (id === ADMIN_ID) return ADMIN.nome;
  const u = lerUsuarios().find(x => x.id === id);
  return u ? u.nome : id;
}
// Quem executou a vistoria (vai pro laudo): nome, título (Eng. ou não) e CREA
function executorDaConta(id) {
  if (!id || id === ADMIN_ID) return { ...ADMIN, ehResponsavelTecnica: true };
  const u = lerUsuarios().find(x => x.id === id);
  if (!u) return { ...ADMIN, ehResponsavelTecnica: true };
  return { id: u.id, nome: u.nome, titulo: u.crea ? 'Eng.' : '', crea: u.crea || '', ehResponsavelTecnica: false };
}

// Proteção simples contra tentativa de adivinhar senha: 8 erros em 15 min bloqueiam o IP por 15 min
const tentativasLogin = new Map();
function ipDe(req) { return String(req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim(); }
function loginBloqueado(ip) {
  const t = tentativasLogin.get(ip);
  return t && t.qtd >= 8 && Date.now() - t.inicio < 15 * 60 * 1000;
}
function registrarFalhaLogin(ip) {
  const t = tentativasLogin.get(ip);
  if (!t || Date.now() - t.inicio > 15 * 60 * 1000) tentativasLogin.set(ip, { qtd: 1, inicio: Date.now() });
  else t.qtd++;
}

const escHtml = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function lerCorpoFormulario(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', c => chunks.push(c));
    req.on('end', () => resolve(new URLSearchParams(Buffer.concat(chunks).toString('utf8'))));
    req.on('error', reject);
  });
}

// ── Páginas simples (login, termo, troca de senha, usuários) ──
const ESTILO_PAGINA = `<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Montserrat',sans-serif;background:#1A3C5E;min-height:100vh;display:flex;align-items:flex-start;justify-content:center;padding:28px 16px;color:#1C1917}
.card{background:#fff;border-radius:16px;padding:28px 24px;max-width:380px;width:100%;box-shadow:0 20px 50px rgba(0,0,0,.25)}
.card.larga{max-width:720px}
.centro{text-align:center}
img.logo{width:64px;height:64px;object-fit:contain;background:#fff;border-radius:12px;padding:6px;border:1px solid #eee;margin-bottom:16px}
h1{font-size:22px;font-weight:800;color:#1A3C5E;margin-bottom:4px}
h2{font-size:15px;font-weight:800;color:#1A3C5E;margin:20px 0 8px}
p.sub{font-size:12px;color:#888;margin-bottom:22px}
label{display:block;font-size:11px;font-weight:700;color:#666;text-transform:uppercase;letter-spacing:.04em;margin:0 0 5px}
input[type=text],input[type=password]{width:100%;padding:12px 13px;border:1.5px solid #ddd;border-radius:9px;font-size:15px;font-family:inherit;margin-bottom:14px;outline:none}
input:focus{border-color:#D4762A}
button,.btn{display:inline-block;width:100%;background:#D4762A;color:#fff;border:none;border-radius:9px;padding:13px;font-size:14px;font-weight:700;font-family:inherit;cursor:pointer;text-align:center;text-decoration:none}
.btn.sec{background:#fff;color:#1A3C5E;border:1.5px solid #1A3C5E}
.btn.peq,button.peq{width:auto;padding:7px 12px;font-size:12px}
.erro{background:#FDECEA;color:#C0392B;font-size:12px;font-weight:600;padding:9px;border-radius:8px;margin-bottom:14px}
.ok{background:#E6F4EC;color:#1A6B3C;font-size:12px;font-weight:600;padding:9px;border-radius:8px;margin-bottom:14px}
.termo{font-size:13px;line-height:1.6;color:#333;max-height:55vh;overflow-y:auto;border:1px solid #eee;border-radius:10px;padding:16px;margin-bottom:16px;background:#FAFAF8}
.termo h3{font-size:13px;margin:14px 0 6px;color:#1A3C5E}
.termo p,.termo li{margin-bottom:8px;text-align:justify}
.termo ol,.termo ul{padding-left:20px}
.check{display:flex;gap:10px;align-items:flex-start;font-size:13px;font-weight:600;margin-bottom:14px;text-transform:none;letter-spacing:0;color:#1C1917}
.check input{width:18px;height:18px;margin-top:2px;flex-shrink:0}
table{width:100%;border-collapse:collapse;font-size:12.5px;margin-bottom:10px}
td,th{padding:9px 8px;border-bottom:1px solid #eee;text-align:left;vertical-align:top}
th{font-size:11px;color:#666;text-transform:uppercase}
.tag{display:inline-block;font-size:10px;font-weight:700;padding:2px 7px;border-radius:20px}
.tag.ok{background:#E6F4EC;color:#1A6B3C;margin:0;padding:2px 7px}.tag.pend{background:#FDECD2;color:#8A5A00}.tag.off{background:#eee;color:#777}
form.inline{display:inline}
.linha{display:flex;gap:8px;flex-wrap:wrap}
.linha > *{flex:1;min-width:140px}
a.voltar{display:inline-block;font-size:13px;font-weight:700;color:#1A3C5E;text-decoration:none;margin-bottom:14px}
</style>`;

function paginaLogin(erro) {
  return `<!DOCTYPE html><html lang="pt-BR"><head>${ESTILO_PAGINA}<title>EngCheck · Login</title></head><body style="align-items:center">
  <form class="card centro" method="POST" action="/login">
    <img class="logo" src="/logo.png" alt="EngCheck">
    <h1>EngCheck</h1>
    <p class="sub">Entre com seu usuário e senha</p>
    ${erro ? `<div class="erro">${escHtml(erro)}</div>` : ''}
    <input type="text" name="usuario" placeholder="Usuário" autocomplete="username" autocapitalize="none" required autofocus>
    <input type="password" name="senha" placeholder="Senha" autocomplete="current-password" required>
    <button type="submit">Entrar</button>
  </form>
</body></html>`;
}

// Termo de Confidencialidade e Proteção de Dados (LGPD). Ao mudar o texto,
// troque TERMO_VERSAO — todas as executoras terão que aceitar de novo.
function textoTermoHtml(nome) {
  return `
<p><strong>TERMO DE CONFIDENCIALIDADE E DE PROTEÇÃO DE DADOS PESSOAIS</strong><br>Sistema EngCheck — G&amp;G Engenharia · Versão ${TERMO_VERSAO}</p>
<p>Pelo presente termo, <strong>${escHtml(nome)}</strong> ("USUÁRIA"), ao receber acesso ao sistema EngCheck, de titularidade e sob responsabilidade técnica da Eng. Civil Gabriela Soares Borges, CREA 427760MG (G&amp;G Engenharia, "CONTROLADORA"), declara estar ciente e de acordo com as condições abaixo.</p>
<h3>1. Objeto</h3>
<p>Este termo regula o acesso e o uso do sistema EngCheck e de todas as informações nele contidas ou obtidas em razão das vistorias realizadas, incluindo dados de clientes, imóveis, fotografias, laudos, agendas, valores e demais informações técnicas, comerciais e financeiras ("INFORMAÇÕES CONFIDENCIAIS").</p>
<h3>2. Dados pessoais tratados</h3>
<p>No uso do sistema, a USUÁRIA terá acesso a dados pessoais de clientes e terceiros, tais como nome, CPF, telefone, endereço, localização e imagens de imóveis, os quais são protegidos pela Lei nº 13.709/2018 (Lei Geral de Proteção de Dados Pessoais — LGPD).</p>
<h3>3. Obrigações da USUÁRIA</h3>
<ol>
<li>Utilizar as INFORMAÇÕES CONFIDENCIAIS e os dados pessoais exclusivamente para a execução das vistorias e a elaboração dos laudos para os quais foi designada, observando as finalidades informadas aos clientes e os princípios do art. 6º da LGPD.</li>
<li>Não divulgar, compartilhar, copiar, encaminhar, publicar ou armazenar fora do sistema quaisquer dados de clientes, fotografias, laudos ou informações do sistema, inclusive em redes sociais, aplicativos de mensagem ou dispositivos pessoais, salvo quando necessário à execução do serviço e com autorização da CONTROLADORA.</li>
<li>Manter sigilo sobre seu usuário e senha, que são pessoais e intransferíveis, não permitindo o uso por terceiros e comunicando imediatamente à CONTROLADORA qualquer suspeita de acesso indevido.</li>
<li>Coletar apenas os dados estritamente necessários à vistoria e registrar fotografias somente do imóvel objeto do serviço, evitando captar pessoas, documentos ou objetos pessoais que não sejam relevantes ao laudo.</li>
<li>Adotar cuidados de segurança no aparelho utilizado (bloqueio de tela, não emprestar o aparelho com o sistema aberto e encerrar a sessão em aparelhos compartilhados).</li>
<li>Comunicar imediatamente à CONTROLADORA qualquer incidente de segurança, perda de aparelho, vazamento ou uso indevido de dados de que tenha conhecimento, para as providências do art. 48 da LGPD.</li>
<li>Não acessar, tentar acessar ou utilizar informações de outras usuárias ou da CONTROLADORA que não lhe tenham sido disponibilizadas.</li>
</ol>
<h3>4. Responsabilidade técnica</h3>
<p>A USUÁRIA declara ciência de que os laudos são emitidos sob responsabilidade técnica da Eng. Civil Gabriela Soares Borges, CREA 427760MG, responsável pela respectiva ART, comprometendo-se a registrar as constatações de forma fiel ao que foi efetivamente observado em campo.</p>
<h3>5. Vigência e devolução</h3>
<p>As obrigações de sigilo e proteção de dados vigoram durante todo o período de acesso ao sistema e permanecem por prazo indeterminado após o seu encerramento. Encerrado o vínculo, a USUÁRIA deverá excluir de seus aparelhos pessoais quaisquer cópias de dados, fotografias e laudos obtidos em razão do serviço, e o acesso ao sistema será desativado.</p>
<h3>6. Descumprimento</h3>
<p>O descumprimento deste termo poderá acarretar o imediato bloqueio do acesso ao sistema, sem prejuízo da responsabilização civil, administrativa e penal cabível, inclusive pelos danos causados à CONTROLADORA, aos clientes e a terceiros, nos termos da legislação vigente e da LGPD.</p>
<h3>7. Registro do aceite</h3>
<p>O aceite eletrônico deste termo, com a digitação do nome completo e a marcação de concordância, será registrado no sistema com data, hora, endereço IP e identificação do aparelho, e tem validade como manifestação de vontade da USUÁRIA.</p>`;
}

function paginaTermo(u, erro) {
  return `<!DOCTYPE html><html lang="pt-BR"><head>${ESTILO_PAGINA}<title>EngCheck · Termo</title></head><body>
  <form class="card larga" method="POST" action="/termo">
    <h1>Termo de confidencialidade</h1>
    <p class="sub">Antes do primeiro acesso, leia e aceite o termo abaixo.</p>
    ${erro ? `<div class="erro">${escHtml(erro)}</div>` : ''}
    <div class="termo">${textoTermoHtml(u.nome)}</div>
    <label>Digite seu nome completo para assinar</label>
    <input type="text" name="nome" placeholder="${escHtml(u.nome)}" required>
    <label class="check"><input type="checkbox" name="aceito" value="1" required> Li, entendi e concordo com o Termo de Confidencialidade e de Proteção de Dados Pessoais.</label>
    <button type="submit">Aceitar e continuar</button>
    <p style="margin-top:12px;text-align:center"><a href="/logout" style="font-size:12px;color:#888">Sair</a></p>
  </form>
</body></html>`;
}

function paginaTrocarSenha(u, erro, obrigatoria) {
  return `<!DOCTYPE html><html lang="pt-BR"><head>${ESTILO_PAGINA}<title>EngCheck · Senha</title></head><body style="align-items:center">
  <form class="card" method="POST" action="/minha-senha">
    ${obrigatoria ? '' : '<a class="voltar" href="/">← Voltar ao app</a>'}
    <h1>${obrigatoria ? 'Crie sua senha' : 'Trocar senha'}</h1>
    <p class="sub">${obrigatoria ? 'Por segurança, troque a senha inicial por uma senha só sua.' : escHtml(u.nome)}</p>
    ${erro ? `<div class="erro">${escHtml(erro)}</div>` : ''}
    <label>Senha atual</label><input type="password" name="atual" autocomplete="current-password" required>
    <label>Nova senha (mín. 8 caracteres, com letras e números)</label><input type="password" name="nova" autocomplete="new-password" required>
    <label>Repita a nova senha</label><input type="password" name="nova2" autocomplete="new-password" required>
    <button type="submit">Salvar nova senha</button>
  </form>
</body></html>`;
}

function paginaUsuarios(msg, erro) {
  const usuarios = lerUsuarios();
  const fmt = (ts) => ts ? new Date(ts).toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' }) : '';
  const linhas = usuarios.map(u => {
    const termo = u.termo && u.termo.versao === TERMO_VERSAO
      ? `<span class="tag ok">Termo aceito</span><br><span style="font-size:11px;color:#666">${fmt(u.termo.em)} · IP ${escHtml(u.termo.ip)}<br>Assinado como "${escHtml(u.termo.nomeDigitado)}"</span>`
      : '<span class="tag pend">Termo pendente</span>';
    const situacao = u.ativo === false ? '<span class="tag off">Desativada</span>' : (u.trocarSenha ? '<span class="tag pend">Senha inicial</span>' : '<span class="tag ok">Ativa</span>');
    return `<tr>
      <td><strong>${escHtml(u.nome)}</strong><br><span style="font-size:11px;color:#666">Usuário: ${escHtml(u.usuario)}${u.crea ? ` · CREA ${escHtml(u.crea)}` : ' · sem CREA'}</span><br>${situacao}</td>
      <td>${termo}</td>
      <td>
        <form class="inline" method="POST" action="/admin/usuarios" onsubmit="return this.senha.value.length>=6 || (alert('A senha provisória precisa ter pelo menos 6 caracteres.'),false)">
          <input type="hidden" name="acao" value="senha"><input type="hidden" name="id" value="${escHtml(u.id)}">
          <input type="text" name="senha" placeholder="Nova senha" style="margin-bottom:6px;padding:7px 9px;font-size:12px">
          <button class="peq" type="submit">Redefinir senha</button>
        </form>
        <form class="inline" method="POST" action="/admin/usuarios" style="margin-top:6px;display:block">
          <input type="hidden" name="acao" value="${u.ativo === false ? 'ativar' : 'desativar'}"><input type="hidden" name="id" value="${escHtml(u.id)}">
          <button class="peq" type="submit" style="background:${u.ativo === false ? '#1A6B3C' : '#C0392B'}" onclick="return confirm('${u.ativo === false ? 'Reativar' : 'Desativar'} o acesso de ${escHtml(u.nome).replace(/'/g, "\\'")}?')">${u.ativo === false ? 'Reativar acesso' : 'Desativar acesso'}</button>
        </form>
      </td></tr>`;
  }).join('');
  return `<!DOCTYPE html><html lang="pt-BR"><head>${ESTILO_PAGINA}<title>EngCheck · Usuários</title></head><body>
  <div class="card larga">
    <a class="voltar" href="/">← Voltar ao app</a>
    <h1>Usuários</h1>
    <p class="sub">Você (Gabriela) é a ADM: entra com o usuário <strong>Gabriela</strong> e sua senha de sempre.</p>
    ${msg ? `<div class="ok">${escHtml(msg)}</div>` : ''}${erro ? `<div class="erro">${escHtml(erro)}</div>` : ''}
    ${usuarios.length ? `<table><tr><th>Executora</th><th>Termo LGPD</th><th>Ações</th></tr>${linhas}</table>` : '<p style="font-size:13px;color:#666;margin-bottom:10px">Nenhuma executora cadastrada ainda.</p>'}
    <h2>Cadastrar executora</h2>
    <form method="POST" action="/admin/usuarios">
      <input type="hidden" name="acao" value="criar">
      <div class="linha">
        <div><label>Usuário (para login)</label><input type="text" name="usuario" placeholder="Ex: Aline" required autocapitalize="none"></div>
        <div><label>Senha inicial</label><input type="text" name="senha" placeholder="Mín. 6 caracteres (ela troca no 1º acesso)" required></div>
      </div>
      <label>Nome completo (sai no laudo)</label><input type="text" name="nome" placeholder="Ex: Aline Baron Camargos de Deus Gomes" required>
      <label>CREA (deixe em branco se não for engenheira)</label><input type="text" name="crea" placeholder="Ex: 353718MG">
      <p style="font-size:11.5px;color:#666;margin:-6px 0 14px">Com CREA, o laudo sai "realizada pela Eng. [nome]". Sem CREA, sai "realizada por [nome]". No primeiro acesso ela aceita o termo e troca a senha.</p>
      <button type="submit">Cadastrar</button>
    </form>
  </div>
</body></html>`;
}

// Trata as rotas de conta. Retorna true se respondeu a requisição.
async function rotasDeConta(req, res, caminho) {
  const html = (status, corpo, extra = {}) => { res.writeHead(status, { 'Content-Type': 'text/html; charset=utf-8', ...extra }); res.end(corpo); return true; };
  const redirecionar = (para, extra = {}) => { res.writeHead(302, { Location: para, ...extra }); res.end(); return true; };

  if (caminho === '/login') {
    if (req.method === 'GET') return html(200, paginaLogin(''));
    if (req.method === 'POST') {
      const ip = ipDe(req);
      if (loginBloqueado(ip)) return html(429, paginaLogin('Muitas tentativas. Aguarde 15 minutos e tente de novo.'));
      const form = await lerCorpoFormulario(req);
      const usuario = normalizarUsuario(form.get('usuario'));
      const senha = form.get('senha') || '';
      if (APP_PASSWORD && ['gabriela', 'admin', 'adm'].includes(usuario) && senha === APP_PASSWORD) {
        return redirecionar('/', { 'Set-Cookie': [cookieSessao(ADMIN_ID, 0), 'engcheck_auth=; Path=/; HttpOnly; Max-Age=0'] });
      }
      const u = lerUsuarios().find(x => normalizarUsuario(x.usuario) === usuario);
      if (u && u.ativo !== false && conferirSenha(senha, u.salt, u.senhaHash)) {
        return redirecionar('/', { 'Set-Cookie': cookieSessao(u.id, u.senhaVersao || 0) });
      }
      registrarFalhaLogin(ip);
      return html(200, paginaLogin(u && u.ativo === false ? 'Este acesso está desativado. Fale com a Gabriela.' : 'Usuário ou senha incorretos.'));
    }
  }

  if (caminho === '/logout') {
    return redirecionar('/login', { 'Set-Cookie': ['engcheck_sessao=; Path=/; HttpOnly; Max-Age=0', 'engcheck_auth=; Path=/; HttpOnly; Max-Age=0', 'engcheck_ver=; Path=/; Max-Age=0'] });
  }

  const u = usuarioDaSessao(req);
  if (!u) return false; // o bloqueio geral (fora daqui) cuida de quem não está logado

  // Sessão antiga da ADM: troca pelo cookie novo, sem pedir login de novo
  if (u.sessaoAntiga && req.method === 'GET' && (req.headers.accept || '').includes('text/html')) {
    res.setHeader('Set-Cookie', [cookieSessao(ADMIN_ID, 0), 'engcheck_auth=; Path=/; HttpOnly; Max-Age=0']);
  }

  if (caminho === '/termo' && u.papel !== 'admin') {
    if (req.method === 'GET') return html(200, paginaTermo(u, ''));
    if (req.method === 'POST') {
      const form = await lerCorpoFormulario(req);
      const nome = (form.get('nome') || '').trim();
      if (form.get('aceito') !== '1') return html(200, paginaTermo(u, 'Marque a caixa de concordância para continuar.'));
      if (normalizarUsuario(nome).replace(/\s+/g, ' ') !== normalizarUsuario(u.nome).replace(/\s+/g, ' ')) return html(200, paginaTermo(u, `Digite seu nome completo exatamente como cadastrado: ${u.nome}`));
      const lista = lerUsuarios();
      const alvo = lista.find(x => x.id === u.id);
      alvo.termo = { versao: TERMO_VERSAO, em: Date.now(), ip: ipDe(req), aparelho: String(req.headers['user-agent'] || '').slice(0, 200), nomeDigitado: nome };
      salvarUsuarios(lista);
      console.log(`✓ Termo de confidencialidade aceito por ${u.nome} (${ipDe(req)})`);
      return redirecionar(alvo.trocarSenha ? '/minha-senha' : '/');
    }
  }

  if (caminho === '/minha-senha' && u.papel !== 'admin') {
    const obrigatoria = !!u.trocarSenha;
    if (req.method === 'GET') return html(200, paginaTrocarSenha(u, '', obrigatoria));
    if (req.method === 'POST') {
      const form = await lerCorpoFormulario(req);
      const lista = lerUsuarios();
      const alvo = lista.find(x => x.id === u.id);
      if (!conferirSenha(form.get('atual') || '', alvo.salt, alvo.senhaHash)) return html(200, paginaTrocarSenha(u, 'Senha atual incorreta.', obrigatoria));
      const nova = form.get('nova') || '';
      if (nova !== (form.get('nova2') || '')) return html(200, paginaTrocarSenha(u, 'As duas senhas novas não são iguais.', obrigatoria));
      if (!senhaForte(nova)) return html(200, paginaTrocarSenha(u, 'A nova senha precisa ter pelo menos 8 caracteres, com letras e números.', obrigatoria));
      if (conferirSenha(nova, alvo.salt, alvo.senhaHash)) return html(200, paginaTrocarSenha(u, 'A nova senha precisa ser diferente da atual.', obrigatoria));
      alvo.salt = crypto.randomBytes(16).toString('hex');
      alvo.senhaHash = hashSenha(nova, alvo.salt);
      alvo.senhaVersao = (alvo.senhaVersao || 0) + 1;
      alvo.trocarSenha = false;
      salvarUsuarios(lista);
      return redirecionar('/', { 'Set-Cookie': cookieSessao(alvo.id, alvo.senhaVersao) });
    }
  }

  if (caminho === '/admin/usuarios' && u.papel === 'admin') {
    if (req.method === 'GET') return html(200, paginaUsuarios('', ''));
    if (req.method === 'POST') {
      const form = await lerCorpoFormulario(req);
      const acao = form.get('acao');
      const lista = lerUsuarios();
      if (acao === 'criar') {
        const usuario = (form.get('usuario') || '').trim();
        const nome = (form.get('nome') || '').trim().replace(/\s+/g, ' ');
        const senha = form.get('senha') || '';
        const crea = (form.get('crea') || '').trim().toUpperCase();
        if (!usuario || !nome) return html(200, paginaUsuarios('', 'Preencha usuário e nome completo.'));
        if (senha.length < 6) return html(200, paginaUsuarios('', 'A senha inicial precisa ter pelo menos 6 caracteres.'));
        const norm = normalizarUsuario(usuario);
        if (['gabriela', 'admin', 'adm'].includes(norm) || lista.some(x => normalizarUsuario(x.usuario) === norm)) return html(200, paginaUsuarios('', `Já existe um usuário "${usuario}".`));
        const salt = crypto.randomBytes(16).toString('hex');
        lista.push({ id: 'u_' + norm.replace(/[^a-z0-9]/g, '') + '_' + crypto.randomBytes(3).toString('hex'), usuario, nome, crea, salt, senhaHash: hashSenha(senha, salt), senhaVersao: 0, trocarSenha: true, ativo: true, criadoEm: Date.now() });
        salvarUsuarios(lista);
        return html(200, paginaUsuarios(`${nome} cadastrada. Usuário: ${usuario}.`, ''));
      }
      const alvo = lista.find(x => x.id === form.get('id'));
      if (!alvo) return html(200, paginaUsuarios('', 'Usuária não encontrada.'));
      if (acao === 'senha') {
        const senha = form.get('senha') || '';
        if (senha.length < 6) return html(200, paginaUsuarios('', 'A senha provisória precisa ter pelo menos 6 caracteres.'));
        alvo.salt = crypto.randomBytes(16).toString('hex');
        alvo.senhaHash = hashSenha(senha, alvo.salt);
        alvo.senhaVersao = (alvo.senhaVersao || 0) + 1; // derruba sessões abertas
        alvo.trocarSenha = true;
        salvarUsuarios(lista);
        return html(200, paginaUsuarios(`Senha de ${alvo.nome} redefinida. No próximo acesso ela cria uma senha nova.`, ''));
      }
      if (acao === 'desativar' || acao === 'ativar') {
        alvo.ativo = acao === 'ativar';
        if (!alvo.ativo) alvo.senhaVersao = (alvo.senhaVersao || 0) + 1; // derruba sessões abertas
        salvarUsuarios(lista);
        return html(200, paginaUsuarios(`Acesso de ${alvo.nome} ${alvo.ativo ? 'reativado' : 'desativado'}.`, ''));
      }
      return html(400, paginaUsuarios('', 'Ação inválida.'));
    }
  }
  return false;
}


// ── HTTP SERVER ──────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  const caminho = req.url.split('?')[0];

  // Login, logout, termo, troca de senha e tela de usuários (ADM)
  try {
    if (await rotasDeConta(req, res, caminho)) return;
  } catch (e) {
    console.error('Erro nas rotas de conta:', e.message);
    res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
    return res.end('Erro interno. Tente de novo.');
  }

  // Bloqueia todo o resto do app pra quem não está logado
  // (ícones/manifest do PWA ficam liberados, senão a própria tela de login não carrega o logo)
  const PUBLICO = ['/logo.png', '/icon-192.png', '/icon-512.png', '/manifest.json', '/service-worker.js'];
  const usuario = usuarioDaSessao(req);
  const querHtml = req.method === 'GET' && (req.headers.accept || '').includes('text/html');
  if (!PUBLICO.includes(caminho)) {
    if (!usuario) {
      if (querHtml) { res.writeHead(302, { Location: '/login' }); return res.end(); }
      res.writeHead(401, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ erro: 'Não autorizado. Faça login novamente.' }));
    }
    // Executora com termo pendente ou senha inicial: só entra depois de resolver
    const pendencia = pendenciaDoUsuario(usuario);
    if (pendencia) {
      const destino = pendencia === 'termo' ? '/termo' : '/minha-senha';
      if (querHtml) { res.writeHead(302, { Location: destino }); return res.end(); }
      res.writeHead(403, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ erro: pendencia === 'termo' ? 'Aceite o termo de confidencialidade para continuar.' : 'Troque sua senha inicial para continuar.', pendencia }));
    }
  }

  // Quem está logado (o app usa pra mostrar o nome e, pra ADM, a guia "Ver conta")
  if (req.method === 'GET' && caminho === '/eu') {
    const ver = contaVisualizada(req, usuario);
    const contas = usuario.papel === 'admin'
      ? [{ id: ADMIN_ID, nome: ADMIN.nome, usuario: ADMIN.usuario }, ...lerUsuarios().map(x => ({ id: x.id, nome: x.nome, usuario: x.usuario, ativo: x.ativo !== false }))]
      : [];
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ id: usuario.id, nome: usuario.nome, usuario: usuario.usuario, papel: usuario.papel, crea: usuario.crea || '', verConta: ver, contas }));
  }

  // ADM escolhe de quem quer ver os dados (Histórico, Agenda, Financeiro, Não finalizadas)
  if (req.method === 'POST' && caminho === '/ver-conta') {
    if (usuario.papel !== 'admin') { res.writeHead(403, { 'Content-Type': 'application/json' }); return res.end(JSON.stringify({ erro: 'Sem permissão.' })); }
    const conta = new URL(req.url, `http://${req.headers.host}`).searchParams.get('conta') || ADMIN_ID;
    const valida = conta === 'todas' || conta === ADMIN_ID || lerUsuarios().some(x => x.id === conta);
    res.writeHead(valida ? 200 : 400, { 'Content-Type': 'application/json', 'Set-Cookie': `engcheck_ver=${encodeURIComponent(valida ? conta : ADMIN_ID)}; Path=/; Secure; SameSite=Lax; Max-Age=31536000` });
    return res.end(JSON.stringify({ ok: valida }));
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
    if (usuario.papel !== 'admin') { res.writeHead(403); return res.end('Sem permissão.'); }
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

  // ── Vistorias não finalizadas: sincronização entre aparelhos ──
  if (caminho.startsWith('/rascunhos')) {
    const urlObj = new URL(req.url, `http://${req.headers.host}`);
    const id = urlObj.searchParams.get('id');
    const responder = (status, obj) => { res.writeHead(status, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(obj)); };
    try {
      // Lista resumida de todas (inclusive as da lixeira e as lápides)
      if (req.method === 'GET' && caminho === '/rascunhos') {
        // escopo=meu: só os do próprio usuário (usado na sincronização do aparelho)
        const soMeus = urlObj.searchParams.get('escopo') === 'meu';
        const daConta = soMeus ? (m) => donoDoItem(m) === usuario.id : filtroDeConta(req, usuario);
        return responder(200, Object.values(lerIndiceRascunhos()).filter(daConta).map(m => ({ ...m, contaNome: nomeDaConta(donoDoItem(m)) })));
      }
      if (!idValido(id)) return responder(400, { erro: 'Identificador inválido.' });
      const metaExistente = lerIndiceRascunhos()[id];
      if (metaExistente && !podeMexer(usuario, metaExistente)) return responder(403, { erro: 'Esta vistoria pertence a outra conta.' });

      // Dados completos de uma vistoria (sem as imagens)
      if (req.method === 'GET' && caminho === '/rascunhos/item') {
        const meta = lerIndiceRascunhos()[id];
        if (!meta || meta.removidoEm) return responder(404, { erro: 'Vistoria não encontrada no servidor.' });
        const estado = JSON.parse(fs.readFileSync(path.join(pastaRascunho(id), 'estado.json'), 'utf8'));
        return responder(200, { ...meta, state: estado });
      }

      // Envia/atualiza uma vistoria. Só aceita se for mais nova que a cópia daqui,
      // assim um aparelho com cópia antiga nunca sobrescreve trabalho mais recente.
      if (req.method === 'PUT' && caminho === '/rascunhos/item') {
        const corpo = await lerCorpoJSON(req);
        if (!corpo || !corpo.state) return responder(400, { erro: 'Dados ausentes.' });
        const idx = lerIndiceRascunhos();
        const atual = idx[id];
        const novo = { atualizadoEm: corpo.atualizadoEm || 0, excluidoEm: corpo.excluidoEm || null, restauradoEm: corpo.restauradoEm || null };
        if (atual && atual.removidoEm && versaoRascunho(novo) <= atual.removidoEm) return responder(200, { ok: false, removido: true, removidoEm: atual.removidoEm });
        if (atual && !atual.removidoEm && versaoRascunho(novo) < versaoRascunho(atual)) return responder(200, { ok: false, maisAntigo: true });
        fs.mkdirSync(pastaFotosRascunho(id), { recursive: true });
        const estadoTmp = path.join(pastaRascunho(id), 'estado.json.tmp');
        fs.writeFileSync(estadoTmp, JSON.stringify(corpo.state));
        fs.renameSync(estadoTmp, path.join(pastaRascunho(id), 'estado.json'));
        // Fotos que não fazem mais parte da vistoria (ex: removidas) saem do disco
        const usadas = new Set(idsFotosDoEstado(corpo.state));
        for (const arq of fs.readdirSync(pastaFotosRascunho(id))) {
          if (!usadas.has(arq.replace(/\.json$/, ''))) { try { fs.unlinkSync(path.join(pastaFotosRascunho(id), arq)); } catch {} }
        }
        const d = corpo.state.dados || {};
        idx[id] = {
          id, contaId: (atual && atual.contaId) || usuario.id, ...novo,
          nome: d.nome || '', tipo: d.tipo || '', locacaoModalidade: d.locacaoModalidade || '',
          qtdRegistros: (corpo.state.registros || []).length, recebidoEm: Date.now()
        };
        salvarIndiceRascunhos(idx);
        const fotosNoServidor = fs.readdirSync(pastaFotosRascunho(id)).map(a => a.replace(/\.json$/, ''));
        return responder(200, { ok: true, fotos: fotosNoServidor });
      }

      // Quais fotos dessa vistoria já estão aqui (pra não reenviar)
      if (req.method === 'GET' && caminho === '/rascunhos/fotos') {
        const pasta = pastaFotosRascunho(id);
        return responder(200, fs.existsSync(pasta) ? fs.readdirSync(pasta).map(a => a.replace(/\.json$/, '')) : []);
      }

      // Recebe um lote de fotos
      if (req.method === 'POST' && caminho === '/rascunhos/fotos') {
        const corpo = await lerCorpoJSON(req);
        const meta = lerIndiceRascunhos()[id];
        if (meta && meta.removidoEm) return responder(200, { ok: false, removido: true });
        fs.mkdirSync(pastaFotosRascunho(id), { recursive: true });
        let salvas = 0;
        for (const f of (corpo.fotos || [])) {
          if (!idValido(f.id) || !f.src) continue;
          fs.writeFileSync(path.join(pastaFotosRascunho(id), f.id + '.json'), JSON.stringify({ src: f.src, thumb: f.thumb || null }));
          salvas++;
        }
        return responder(200, { ok: true, salvas });
      }

      // Devolve um lote de fotos (pra abrir a vistoria em outro aparelho)
      if (req.method === 'POST' && caminho === '/rascunhos/fotos/baixar') {
        const corpo = await lerCorpoJSON(req);
        const fotos = {};
        for (const fid of (corpo.ids || [])) {
          if (!idValido(fid)) continue;
          const arq = path.join(pastaFotosRascunho(id), fid + '.json');
          if (fs.existsSync(arq)) fotos[fid] = JSON.parse(fs.readFileSync(arq, 'utf8'));
        }
        return responder(200, { fotos });
      }

      // Mover pra lixeira / restaurar
      if (req.method === 'POST' && caminho === '/rascunhos/acao') {
        const acao = urlObj.searchParams.get('acao');
        const idx = lerIndiceRascunhos();
        const m = idx[id];
        if (!m || m.removidoEm) return responder(200, { ok: false, naoEncontrado: true });
        if (acao === 'excluir') m.excluidoEm = Date.now();
        else if (acao === 'restaurar') { m.excluidoEm = null; m.restauradoEm = Date.now(); }
        else return responder(400, { erro: 'Ação inválida.' });
        salvarIndiceRascunhos(idx);
        return responder(200, { ok: true, ...m });
      }

      // Apaga de vez (vistoria finalizada com laudo gerado, descartada ou apagada da lixeira)
      if (req.method === 'DELETE' && caminho === '/rascunhos/item') {
        const idx = lerIndiceRascunhos();
        apagarPastaRascunho(id);
        idx[id] = { id, contaId: (idx[id] && idx[id].contaId) || usuario.id, removidoEm: Date.now() };
        salvarIndiceRascunhos(idx);
        return responder(200, { ok: true });
      }
      return responder(404, { erro: 'Rota não encontrada.' });
    } catch (e) {
      console.error('Erro em', caminho, e.message);
      return responder(500, { erro: e.message });
    }
  }

  // ── Agenda: pré-cadastro dos dados do cliente/imóvel antes da vistoria ──
  if (req.method === 'GET' && req.url === '/agenda') {
    const lista = lerAgenda().filter(filtroDeConta(req, usuario)).sort((a, b) => (a.dados.data || '9999').localeCompare(b.dados.data || '9999'));
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify(lista));
  }

  if (req.method === 'POST' && req.url === '/agenda') {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      try {
        const payload = JSON.parse(Buffer.concat(chunks).toString('utf8'));
        if (!payload.dados || !payload.dados.nome) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          return res.end(JSON.stringify({ erro: 'Nome do cliente é obrigatório.' }));
        }
        const agenda = lerAgenda();
        const item = {
          id: 'ag_' + Date.now() + Math.random().toString(36).slice(2),
          contaId: usuario.id,
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
    const agenda = lerAgenda().filter(a => !(a.id === id && podeMexer(usuario, a)));
    salvarAgenda(agenda);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ ok: true }));
  }

  if (req.method === 'PUT' && req.url.startsWith('/agenda')) {
    const urlObj = new URL(req.url, `http://${req.headers.host}`);
    const id = urlObj.searchParams.get('id');
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      try {
        const payload = JSON.parse(Buffer.concat(chunks).toString('utf8'));
        const agenda = lerAgenda();
        const item = agenda.find(a => a.id === id && podeMexer(usuario, a));
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

  if (req.method === 'GET' && req.url.startsWith('/historico')) {
    const urlObj = new URL(req.url, `http://${req.headers.host}`);
    const verLixeira = urlObj.searchParams.get('lixeira') === '1';
    const hist = lerHistorico();
    const daConta = filtroDeConta(req, usuario);
    const filtrado = (verLixeira ? hist.filter(h => h.excluidoEm) : hist.filter(h => !h.excluidoEm)).filter(daConta)
      .map(h => ({ ...h, contaNome: nomeDaConta(donoDoItem(h)) }));
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify(filtrado.reverse()));
  }

  // Restaura um laudo que estava na lixeira, tirando ele de lá — volta a
  // aparecer no Histórico normal, como se nunca tivesse sido excluído.
  if (req.method === 'POST' && req.url.startsWith('/historico/restaurar')) {
    try {
      const urlObj = new URL(req.url, `http://${req.headers.host}`);
      const arquivoAlvo = urlObj.searchParams.get('arquivo');
      const hist = lerHistorico();
      const item = hist.find(h => h.arquivo === arquivoAlvo && podeMexer(usuario, h));
      if (!item) throw new Error('Laudo não encontrado.');
      delete item.excluidoEm;
      salvarHistorico(hist);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ ok: true }));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ erro: e.message }));
    }
  }

  // "Excluir" agora é uma exclusão SUAVE — o laudo some do Histórico normal e
  // vai pra uma lixeira, mas o arquivo continua guardado no servidor por 30
  // dias, dando uma segunda chance em caso de exclusão sem querer. Só depois
  // desses 30 dias (ou se a pessoa esvaziar a lixeira manualmente) o arquivo é
  // apagado de vez, sem volta.
  // Deixa editar o valor e/ou a forma de pagamento de um laudo já concluído —
  // útil porque muitas vezes só se sabe como o cliente vai pagar no dia da
  // vistoria (ou até depois), e antes não dava pra corrigir isso no Financeiro.
  if (req.method === 'PUT' && req.url.startsWith('/historico')) {
    try {
      const urlObj = new URL(req.url, `http://${req.headers.host}`);
      const arquivoAlvo = urlObj.searchParams.get('arquivo');
      const chunks = [];
      req.on('data', chunk => chunks.push(chunk));
      req.on('end', () => {
        try {
          const payload = JSON.parse(Buffer.concat(chunks).toString('utf8'));
          const hist = lerHistorico();
          const item = hist.find(h => h.arquivo === arquivoAlvo && podeMexer(usuario, h));
          if (!item) throw new Error('Laudo não encontrado.');
          if (payload.valor !== undefined) item.valor = payload.valor;
          if (payload.formaPagamento !== undefined) item.formaPagamento = payload.formaPagamento;
          salvarHistorico(hist);
          res.writeHead(200, { 'Content-Type': 'application/json' });
          return res.end(JSON.stringify({ ok: true }));
        } catch (e) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          return res.end(JSON.stringify({ erro: e.message }));
        }
      });
      return;
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ erro: e.message }));
    }
  }

  if (req.method === 'DELETE' && req.url.startsWith('/historico')) {
    try {
      const urlObj = new URL(req.url, `http://${req.headers.host}`);
      const arquivoAlvo = urlObj.searchParams.get('arquivo'); // se vier, mexe só nesse; senão, em tudo
      const permanente = urlObj.searchParams.get('permanente') === '1'; // apaga de vez, sem passar pela lixeira
      const hist = lerHistorico();
      // Sem "arquivo": mexe só no que pertence à conta visualizada (nunca no de outra pessoa)
      const alvos = arquivoAlvo ? hist.filter(h => h.arquivo === arquivoAlvo && podeMexer(usuario, h)) : hist.filter(filtroDeConta(req, usuario));

      if (permanente) {
        for (const item of alvos) {
          const docxPath = path.join(OUTPUT_DIR, item.arquivo);
          const pdfPath = docxPath.replace('.docx', '.pdf');
          [docxPath, docxPath + '.gz', pdfPath, pdfPath + '.gz'].forEach(p => {
            try { if (fs.existsSync(p)) fs.unlinkSync(p); } catch (e) { /* segue mesmo se algum arquivo já não existir */ }
          });
        }
        const alvosArquivos = new Set(alvos.map(a => a.arquivo));
        salvarHistorico(hist.filter(h => !alvosArquivos.has(h.arquivo)));
      } else {
        const agora = Date.now();
        alvos.forEach(item => { item.excluidoEm = agora; });
        salvarHistorico(hist);
      }

      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ ok: true, removidos: alvos.length, permanente }));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ erro: e.message }));
    }
  }

  if (req.method === 'GET' && req.url.startsWith('/download-pdf/')) {
    const filename = decodeURIComponent(req.url.replace('/download-pdf/', ''));
    const registroLaudo = lerHistorico().find(h => h.arquivo === filename);
    if (filename.includes('..') || filename.includes('/') || (registroLaudo ? !podeMexer(usuario, registroLaudo) : usuario.papel !== 'admin')) {
      res.writeHead(404); return res.end('Not found');
    }
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
    const registroLaudo = lerHistorico().find(h => h.arquivo === filename);
    if (filename.includes('..') || filename.includes('/') || (registroLaudo ? !podeMexer(usuario, registroLaudo) : usuario.papel !== 'admin')) {
      res.writeHead(404); return res.end('Not found');
    }
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
      const { registros, tipoVistoria, obsGeral, dados } = await lerCorpoJSON(req);
      const ehLocacao = tipoVistoria === 'Locação';
      if (tipoVistoria === 'Constatação') {
        const texto = await gerarConclusaoConstatacao({ registros, obsGeral, dados: dados || {} });
        res.writeHead(200, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ conclusao: texto.trim() }));
      }
      const texto = await gerarConclusaoVistoria({ registros, obsGeral, tipoVistoria, dados: dados || {} });
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ conclusao: texto.trim() }));
    } catch (e) {
      console.error('Erro em /gerar-conclusao:', e.message);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ conclusao: '' })); // falha silenciosa: server usa o texto padrão fixo
    }
  }

  // Laudo de Constatação: transforma o relato informal do cliente (anotado às pressas
  // na vistoria) num texto de ANAMNESE técnico e formal, no estilo dos laudos da G&G.
  if (req.method === 'POST' && req.url === '/gerar-anamnese-texto') {
    try {
      const { relato, nome } = await lerCorpoJSON(req);
      if (!relato || !relato.trim()) throw new Error('Relato vazio.');
      const texto = await chamarClaudeTecnico({
        maxTokens: 1500,
        system: 'Você é um(a) engenheiro(a) civil redigindo a seção ANAMNESE de um Laudo de Constatação de imóvel. Vai receber anotações informais com o relato do(a) contratante sobre o histórico do imóvel e dos problemas. Reescreva em 2 a 5 parágrafos técnicos, formais e em terceira pessoa, iniciando com algo como "Conforme informações prestadas pelo(a) contratante..." ou "Segundo relato...", deixando claro que se trata de informação relatada (não constatada). Organize em ordem cronológica: motivo da solicitação, problemas originais, intervenções realizadas (por quem, quando), o que surgiu/piorou depois e as dúvidas/preocupações do(a) contratante. NÃO invente fatos, datas, nomes, quantidades ou causas que não estejam no relato — use apenas o que foi informado. Termine com um parágrafo dizendo que, diante desse cenário, foi solicitada a elaboração do presente Laudo de Constatação para registrar, por inspeção exclusivamente visual, as condições aparentes da edificação, sem ensaios, testes, cálculos estruturais, medições instrumentais ou procedimentos invasivos. Retorne SOMENTE os parágrafos, separados por uma linha em branco, sem título, sem markdown, sem numeração.',
        messages: [{ role: 'user', content: `Contratante: ${nome || 'não informado'}\n\nAnotações do relato:\n${relato.trim()}` }]
      });
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ texto: texto.trim() }));
    } catch (e) {
      console.error('Erro em /gerar-anamnese-texto:', e.message);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ erro: e.message }));
    }
  }

  // Transforma uma descrição curta/informal de defeito (ex: "buraco na parede") no
  // texto técnico formal que vai pro laudo, no mesmo estilo da biblioteca de legendas.
  if (req.method === 'POST' && req.url === '/gerar-defeito-texto') {
    try {
      const { ambiente, textoCurto } = await lerCorpoJSON(req);
      if (!textoCurto || !textoCurto.trim()) throw new Error('Texto do defeito vazio.');
      const texto = await chamarClaude({
        model: 'claude-haiku-4-5-20251001',
        maxTokens: 200,
        system: 'Você é um(a) engenheiro(a) civil redigindo a descrição técnica de UM defeito construtivo específico, pra entrar na seção de registros fotográficos de um laudo de vistoria de imóvel. Use o mesmo estilo formal e objetivo de normas ABNT e do IBAPE Nacional. Escreva de 1 a 2 frases apenas, descrevendo o defeito de forma técnica, e se fizer sentido, uma recomendação breve de correção. NÃO use saudações, títulos, aspas, markdown ou numeração — retorne SOMENTE o texto final, pronto pra ser colado no laudo, como neste exemplo de estilo: "Foram identificados danos na superfície da parede, comprometendo a integridade e o acabamento do revestimento. Recomenda-se reparo com massa adequada e repintura do trecho afetado."' + '\n\n' + REGRA_NORMAS,
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

  // Registra um serviço concluído SEM gerar laudo novo — pra revistorias em que
  // não é necessário produzir um documento (ex: já foi tudo corrigido, só
  // precisa lançar no Financeiro). Fica no Histórico sem arquivo pra baixar.
  if (req.method === 'POST' && req.url === '/finalizar-sem-laudo') {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      try {
        const payload = JSON.parse(Buffer.concat(chunks).toString('utf8'));
        if (!payload.dados || !payload.dados.nome) throw new Error('Nome do cliente é obrigatório.');
        const hist = lerHistorico();
        hist.push({
          contaId: usuario.id,
          executor: executorDaConta(usuario.id).nome,
          nome: payload.dados.nome,
          empreendimento: payload.dados.empreendimento || '',
          endereco: payload.dados.endereco || '',
          data: payload.dados.data || '—',
          tipo: payload.tipoVistoria || 'Revistoria',
          valor: payload.dados.valor || null,
          formaPagamento: payload.dados.formaPagamento || null,
          registros: [],
          arquivo: `semdoc_${Date.now()}${Math.random().toString(36).slice(2, 8)}`, // não existe arquivo de verdade, é só um identificador único
          semDocumento: true,
          ts: Date.now()
        });
        salvarHistorico(hist);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ ok: true }));
      } catch (e) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ error: e.message }));
      }
    });
    return;
  }

  if (req.method === 'POST' && req.url === '/gerar-laudo') {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', async () => {
      try {
        const payload = JSON.parse(Buffer.concat(chunks).toString('utf8'));
        const executor = executorDaConta(usuario.id);
        const docxBuffer = await gerarDocx(payload, executor);
        const ts = Date.now();
        const nomeArquivo = `Laudo_${(payload.dados.nome||'cliente').replace(/\s+/g,'_')}_${payload.dados.data||'vistoria'}_${ts}.docx`;
        const docxPath = path.join(OUTPUT_DIR, nomeArquivo);
        fs.writeFileSync(docxPath, docxBuffer);

        const hist = lerHistorico();
        const registrosResumo = (payload.registros || []).map(r => ({ ambiente: r.ambiente, defeito: r.defeito }));
        hist.push({ contaId: usuario.id, executor: executor.nome, nome: payload.dados.nome||'Cliente', empreendimento: payload.dados.empreendimento||'', endereco: payload.dados.endereco||'', data: payload.dados.data||'—', tipo: payload.tipoVistoria||'—', modalidade: payload.dados.locacaoModalidade || null, valor: payload.dados.valor || null, formaPagamento: payload.dados.formaPagamento || null, registros: registrosResumo, arquivo: nomeArquivo, ts });
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
