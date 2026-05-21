// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║          PRATTI COMERCIAL — TEAMS CRM BOT  |  v3.0 FORTRESS EDITION         ║
// ║     Arquitetura: Engenharia Defensiva Extrema + NLP Omnisciente              ║
// ║     Conformidade: Regras 1-9 do Desafio 3 — Hackathon Biopark                ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

'use strict';

require('dotenv').config();
const { App } = require('@microsoft/teams.apps');

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// DIRETIVA 1 — ANTI-CRASH GLOBAL: captura qualquer exceção não tratada antes
// que o processo Node.js seja derrubado. NADA escapa.
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
process.on('uncaughtException', (err) => {
  registrarLog(
    'FATAL',
    'SISTEMA',
    'UNCAUGHT_EXCEPTION',
    { stack: err.stack },
    'Exceção não tratada capturada no nível do processo — bot continua ativo.'
  );
});

process.on('unhandledRejection', (reason) => {
  registrarLog(
    'FATAL',
    'SISTEMA',
    'UNHANDLED_REJECTION',
    { reason: String(reason) },
    'Promise rejeitada sem handler — isolada pelo watchdog global.'
  );
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// IDENTIDADE DO CORRETOR LOGADO (simulado para o Hackathon)
// Em produção, estes valores viriam da sessão autenticada do Teams/CV CRM.
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const CORRETOR_NOME = process.env.CORRETOR_NOME || 'Leonardo';
const CORRETOR_ID = parseInt(process.env.CORRETOR_ID || '464', 10);

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// STORAGE IN-MEMORY (compatível com qualquer ambiente sem dependência externa)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const storage = {
  read: async () => ({}),
  write: async () => {},
  delete: async () => {},
};

const app = new App({
  storage,
  config: {
    MicrosoftAppId: process.env.MICROSOFT_APP_ID || '',
    MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD || '',
  },
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// BANCO LOCAL — cache sincronizado para validações de governança offline
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
let bancoDadosCRM = [
  {
    id: 101,
    nome: 'João Silva',
    telefone: '45999991111',
    email: 'joao.silva101@teste.com',
    empreendimento: 'Carmel',
    responsavel: 'Alex Wilber',
    corretorId: 120,
  },
  {
    id: 102,
    nome: 'João Silva',
    telefone: '45988882222',
    email: 'joao.silva102@teste.com',
    empreendimento: 'Villa Bella',
    responsavel: 'Alex Wilber',
    corretorId: 120,
  },
  {
    id: 103,
    nome: 'Carlos Souza',
    telefone: '45977773333',
    email: 'carlos.souza@teste.com',
    empreendimento: 'Carmel',
    responsavel: 'Mariana Costa',
    corretorId: 200,
  },
  {
    id: 464,
    nome: 'Pedro Almada',
    telefone: '45966664444',
    email: 'pedro.almada@teste.com',
    empreendimento: 'Villa Bella',
    responsavel: CORRETOR_NOME,
    corretorId: CORRETOR_ID,
  },
  {
    id: 465,
    nome: 'Ana Lima',
    telefone: '45955553333',
    email: 'ana.lima@teste.com',
    empreendimento: 'Carmel',
    responsavel: CORRETOR_NOME,
    corretorId: CORRETOR_ID,
  },
];

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// MÁQUINA DE ESTADOS DE CONVERSA — com timestamp de expiração (60 segundos)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const memoriaConversa = {};
const EXPIRACAO_ESTADO_MS = 60_000; // 60 segundos

function obterEstado(userId) {
  const estado = memoriaConversa[userId];
  if (!estado) return null;
  if (Date.now() - estado.timestamp > EXPIRACAO_ESTADO_MS) {
    delete memoriaConversa[userId];
    registrarLog(
      'INFO',
      userId,
      'ESTADO_EXPIRADO',
      {},
      'Estado de conversa expirado automaticamente por inatividade (60s).'
    );
    return null;
  }
  return estado;
}

function definirEstado(userId, dados) {
  memoriaConversa[userId] = { ...dados, timestamp: Date.now() };
}

function limparEstado(userId) {
  delete memoriaConversa[userId];
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// DIRETIVA 9 — LOGGER DE AUDITORIA ULTRA-DETALHADO
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function registrarLog(tipo, usuario, evento, payload = {}, decisao = '') {
  const ts = new Date().toISOString();
  const borda =
    tipo === 'VIOLAÇÃO' || tipo === 'FATAL'
      ? '🚨 ════════════════════════════════════════════════════════════ 🚨'
      : '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━';

  const icone =
    {
      SUCESSO: '✅',
      ERRO: '❌',
      BLOQUEIO: '🔒',
      VIOLAÇÃO: '🚨',
      INFO: '📋',
      FATAL: '💀',
      WARN: '⚠️ ',
    }[tipo] || '🔵';

  console.log(`\n${borda}`);
  console.log(`${icone}  [${tipo}] — ${ts}`);
  console.log(`   👤 Usuário Ativo : ${usuario}`);
  console.log(`   🎯 Evento        : ${evento}`);
  if (Object.keys(payload).length > 0) {
    console.log(
      `   📦 Payload Limpo : ${JSON.stringify(payload, null, 2).replace(/\n/g, '\n               ')}`
    );
  }
  if (decisao) {
    console.log(`   🧠 [DECISÃO]     : ${decisao}`);
  }
  console.log(`${borda}\n`);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// DIRETIVA 6 — CIRCUIT BREAKER: AbortController com timeout de 6 segundos
// Trata HTTP 400, 404, 500, JSON malformado e rede inativa.
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async function enviarParaCRM(metodo, endpoint, payload = null) {
  const urlApi = `${process.env.CRM_API_URL || 'https://mock.cvcrm.local'}${endpoint}`;
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 6_000);

  try {
    const config = {
      method: metodo,
      signal: controller.signal,
      headers: {
        'Content-Type': 'application/json',
        token: process.env.CRM_TOKEN || '',
        email: process.env.CRM_EMAIL || '',
      },
    };

    if (payload && ['POST', 'PUT', 'PATCH'].includes(metodo.toUpperCase())) {
      config.body = JSON.stringify(payload);
    }

    registrarLog(
      'INFO',
      'SISTEMA',
      `HTTP_${metodo.toUpperCase()}`,
      { url: urlApi, payload },
      `Requisição enviada ao CRM.`
    );

    const resposta = await fetch(urlApi, config);
    clearTimeout(timeoutId);

    // Parse seguro do JSON
    let dadosRetorno = null;
    try {
      const textoResposta = await resposta.text();
      dadosRetorno = textoResposta ? JSON.parse(textoResposta) : {};
    } catch {
      dadosRetorno = { message: 'Resposta da API não é um JSON válido.' };
    }

    if (resposta.ok) {
      return { sucesso: true, dados: dadosRetorno, status: resposta.status };
    }

    const msgErro =
      dadosRetorno?.message ||
      dadosRetorno?.erro ||
      dadosRetorno?.error ||
      `Erro HTTP ${resposta.status}`;

    if (resposta.status === 404)
      return { sucesso: false, erro: `Lead não localizado no CRM (404).` };
    if (resposta.status === 400)
      return { sucesso: false, erro: `Dados inválidos rejeitados pelo CRM: ${msgErro}` };
    if (resposta.status === 401 || resposta.status === 403)
      return {
        sucesso: false,
        erro: `Credenciais inválidas ou acesso negado pelo CRM (${resposta.status}).`,
      };
    if (resposta.status >= 500)
      return {
        sucesso: false,
        erro: `Servidor do CRM está com instabilidade (${resposta.status}). Tente novamente em instantes.`,
      };

    return { sucesso: false, erro: msgErro };
  } catch (err) {
    clearTimeout(timeoutId);
    if (err.name === 'AbortError') {
      registrarLog(
        'ERRO',
        'SISTEMA',
        'CRM_TIMEOUT',
        {},
        'Requisição abortada após 6 segundos — Circuit Breaker ativado.'
      );
      return {
        sucesso: false,
        erro: '⏱️ O servidor do CRM demorou mais de 6 segundos para responder. O serviço pode estar sobrecarregado. Tente novamente.',
      };
    }
    registrarLog(
      'ERRO',
      'SISTEMA',
      'CRM_CONEXAO_FALHA',
      { erro: err.message },
      'Falha de rede ao tentar alcançar o CRM.'
    );
    return { sucesso: false, erro: `Falha de conexão com o CRM: ${err.message}` };
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// DIRETIVA 3 — HELPERS DE REGEX BULLETPROOF
// Toda extração de padrão é feita por estas funções seguras. Zero TypeErrors.
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const REGEX_TELEFONE = /(?:\(?\d{2}\)?\s?)?\d{4,5}[-\s]?\d{4}/g;
const REGEX_EMAIL = /\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b/gi;
const REGEX_ID_LEAD = /\b(?:lead|id|código|codigo|numero|nº|n°)?\s*#?(\d{2,6})\b/gi;
const REGEX_NOME_CLIENTE =
  /(?:cliente|lead|contato|para\s+o|de|do)?\s+([A-ZÀ-Ÿa-zà-ÿ]{2,}(?:\s+[A-ZÀ-Ÿa-zà-ÿ]{2,}){1,4})/i;

function extrairTelefone(texto) {
  try {
    const resultado = texto.match(REGEX_TELEFONE);
    return Array.isArray(resultado) && resultado.length > 0 ? resultado[0].trim() : null;
  } catch {
    return null;
  }
}

function extrairEmail(texto) {
  try {
    const resultado = texto.match(REGEX_EMAIL);
    return Array.isArray(resultado) && resultado.length > 0
      ? resultado[0].trim().toLowerCase()
      : null;
  } catch {
    return null;
  }
}

function extrairIds(texto) {
  try {
    // Remove telefones do texto antes de buscar IDs numéricos
    const textoSemFone = texto.replace(REGEX_TELEFONE, ' ');
    const ids = [];
    let m;
    const re = /\b(\d{2,6})\b/g;
    while ((m = re.exec(textoSemFone)) !== null) {
      ids.push(parseInt(m[1], 10));
    }
    return ids.length > 0 ? ids : null;
  } catch {
    return null;
  }
}

function extrairNome(texto) {
  try {
    // Remove telefones, emails e IDs do texto para isolar o nome
    const limpo = texto
      .replace(REGEX_TELEFONE, ' ')
      .replace(REGEX_EMAIL, ' ')
      .replace(/\b\d{2,6}\b/g, ' ')
      .replace(
        /\b(crie|criar|cadastre|cadastrar|novo lead|novo cliente|adicione|adicionar|inclua|incluir|insira|inserir|registre|registrar|lead|cliente|contato|fone|telefone|email|empreendimento|obs|observa[cç][aã]o|nota|corretor|para|com|de|do|da|no|na|e|o|a)\b/gi,
        ' '
      )
      .replace(/[!@#$%^&*()_+=\[\]{};':"\\|,.<>?/`~\-]/g, ' ')
      .replace(/\s{2,}/g, ' ')
      .trim();

    const palavras = limpo.split(' ').filter((p) => p.length >= 2 && /^[A-ZÀ-Ÿa-zà-ÿ]/.test(p));
    if (palavras.length === 0) return null;

    // Retorna até 3 palavras seguidas que pareçam um nome próprio
    const candidatos = [];
    for (let i = 0; i < Math.min(palavras.length, 5); i++) {
      if (/^[A-ZÀ-Ÿ]/.test(palavras[i])) {
        candidatos.push(palavras[i]);
        if (candidatos.length === 3) break;
      } else if (candidatos.length > 0) break;
    }
    return candidatos.length > 0 ? candidatos.join(' ') : palavras.slice(0, 2).join(' ');
  } catch {
    return null;
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// EXTRATORES DE CAMPOS OPCIONAIS — enriquecem o lead além do mínimo obrigatório
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const REGEX_CPF = /\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b/g;
const REGEX_RENDA = /renda\s+(?:aproximada\s+|mensal\s+)?(?:de\s+)?R?\$?\s*(\d[\d.,]+)\s*(mil)?/i;

function extrairOrigem(texto) {
  try {
    const match = texto.match(
      /(?:origem|veio|captado|captei|captou|indicado|indicação|indicacao|canal|vindo)\s*(?::|do|da|de|via|pelo|pela|no|na|através de|atraves de)?\s*([A-ZÀ-Ÿa-zà-ÿ][A-ZÀ-Ÿa-zà-ÿ\s]{2,25})/i
    );
    if (match) return match[1].trim();
    const origens = [
      ['instagram', 'Instagram'],
      ['facebook', 'Facebook'],
      ['whatsapp', 'WhatsApp'],
      ['site', 'Site'],
      ['portal', 'Portal'],
      ['indicação', 'Indicação'],
      ['indicacao', 'Indicação'],
      ['google', 'Google Ads'],
      ['outdoor', 'Outdoor'],
      ['feirão', 'Feirão'],
      ['feirao', 'Feirão'],
      ['plantão', 'Plantão'],
      ['plantao', 'Plantão'],
      ['evento', 'Evento'],
      ['stand', 'Stand'],
      ['tiktok', 'TikTok'],
      ['youtube', 'YouTube'],
      ['email mkt', 'E-mail Marketing'],
      ['cold call', 'Cold Call'],
      ['telefone', 'Ligação'],
      ['visita', 'Visita Espontânea'],
    ];
    const t = texto.toLowerCase();
    for (const [chave, label] of origens) {
      if (t.includes(chave)) return label;
    }
    return null;
  } catch {
    return null;
  }
}

function extrairRenda(texto) {
  try {
    const match = texto.match(REGEX_RENDA);
    if (!match) return null;
    let valor = parseFloat(match[1].replace(/\./g, '').replace(',', '.'));
    if (match[2] && match[2].toLowerCase() === 'mil') valor *= 1000;
    return isNaN(valor) ? null : `R$ ${valor.toLocaleString('pt-BR')}`;
  } catch {
    return null;
  }
}

function extrairCPF(texto) {
  try {
    const match = texto.match(REGEX_CPF);
    return match ? match[0] : null;
  } catch {
    return null;
  }
}

function extrairDataRetorno(texto) {
  try {
    // Captura padrões como "retorno amanhã às 14h", "ligar na sexta às 10h", "follow-up dia 25/06"
    const match = texto.match(
      /(?:retorno|ligar|callback|follow[- ]?up|agendar?|chamar|contatar|marcar)\s+(?:para\s+)?(?:o\s+dia\s+|dia\s+)?(\d{1,2}\/\d{1,2}(?:\/\d{2,4})?|amanhã|amanha|hoje|segunda(?:-feira)?|terça(?:-feira)?|terca(?:-feira)?|quarta(?:-feira)?|quinta(?:-feira)?|sexta(?:-feira)?|sábado|sabado|domingo)(?:\s+(?:às|as|@)\s*(\d{1,2}[h:]\d{0,2}))?/i
    );
    if (!match) return null;
    const dia = match[1];
    const hora = match[2] ? match[2].replace('h', ':00').replace(/(\d+):$/, '$1:00') : null;
    return hora ? `${dia} às ${hora}` : dia;
  } catch {
    return null;
  }
}

function extrairObservacaoInicial(texto) {
  try {
    const match = texto.match(
      /(?:obs(?:ervação|ervacao)?|nota|observação|observacao|detalhe|interesse|comentário|comentario)[:\s]+(.{5,120})/i
    );
    return match ? match[1].trim() : null;
  } catch {
    return null;
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// DIRETIVA 4 — SANEAMENTO DE INPUT
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const LIMITE_CARACTERES = 600;

function sanitizarInput(texto) {
  if (typeof texto !== 'string') return '';
  // Remove tags HTML/XML que o Teams às vezes injeta
  let limpo = texto.replace(/<[^>]*>/g, '').trim();
  // Reduz repetições absurdas de caracteres especiais (ex: "!!!!!!!!!")
  limpo = limpo.replace(/([^a-zA-Z0-9À-ÿ\s@.\-_])\1{3,}/g, '$1');
  // Trunca se necessário
  if (limpo.length > LIMITE_CARACTERES) {
    limpo = limpo.substring(0, LIMITE_CARACTERES);
  }
  return limpo;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// DIRETIVA 2 — NLP OMNISCIENTE: DICIONÁRIOS COLOSSAIS DE INTENÇÃO
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const INTENCOES = {
  CRIAR: [
    'crie',
    'criar',
    'criando',
    'cria',
    'cadastre',
    'cadastrar',
    'cadastrando',
    'cadastra',
    'cadastro',
    'novo lead',
    'novo cliente',
    'nova entrada',
    'novo contato',
    'adicione',
    'adicionar',
    'adicionando',
    'adiciona',
    'inclua',
    'incluir',
    'incluindo',
    'inclui',
    'insira',
    'inserir',
    'inserindo',
    'insere',
    'registrar cliente',
    'registre o cliente',
    'registre cliente',
    'nova prospecção',
    'nova prospeccao',
    'prospecção nova',
    'subir lead',
    'subir cliente',
    'sobe lead',
    'botar no sistema',
    'colocar no sistema',
    'lançar lead',
    'lançar cliente',
    'lancar lead',
    'lancar cliente',
    'dar entrada',
    'dar entrada no lead',
    'entrada de lead',
    'salvar contato',
    'salvar cliente',
    'salva contato',
    'salva cliente',
    'fazer cadastro',
    'faça cadastro',
    'faz cadastro',
    'abrir lead',
    'abre lead',
    'abre cliente',
    'novo prospect',
    'novo interessado',
    'nova captação',
    'nova captacao',
    'captou lead',
    'captei lead',
    'capturar lead',
    'gerar lead',
    'gera lead',
    'loga cliente',
    'logar cliente',
    'coloca lead',
    'colocar lead',
    'registrar lead',
    'registre lead',
    'add lead',
    'add cliente',
    'novo reg',
    'novo registro',
    'criar contato',
    'cria contato',
    'cadastrar contato',
    'nova ficha',
    'abre ficha',
    'abrir ficha',
    'criar ficha',
    'cria ficha',
  ],

  LISTAR: [
    'liste',
    'listar',
    'listando',
    'lista',
    'mostre',
    'mostrar',
    'mostrando',
    'mostra',
    'exiba',
    'exibir',
    'exibindo',
    'exibe',
    'ver todos',
    'veja todos',
    'ver leads',
    'veja leads',
    'consultar',
    'consulta',
    'consultando',
    'consulte',
    'puxar',
    'puxa',
    'puxa leads',
    'puxar leads',
    'relatório',
    'relatorio',
    'gera relatório',
    'gera relatorio',
    'ver relatorio',
    'ver relatório',
    'meus leads',
    'meus clientes',
    'minha carteira',
    'carteira de leads',
    'cadê os leads',
    'cade os leads',
    'cadê meus leads',
    'buscar leads',
    'busca leads',
    'buscar clientes',
    'trazer base',
    'traz base',
    'trazer leads',
    'traz leads',
    'trazer carteira',
    'todos os leads',
    'todos os clientes',
    'ver base',
    'mostra base',
    'base de leads',
    'listar clientes',
    'ver clientes',
    'quais leads',
    'quais clientes',
    'que leads',
    'que clientes',
    'lista de leads',
    'lista de clientes',
    'show leads',
    'mostrar base',
    'exibe base',
    'exibir base',
    'ver registros',
    'listar registros',
    'meus registros',
    'meus contatos',
    'ver contatos',
    'trazer contatos',
  ],

  ATUALIZAR: [
    'atualize',
    'atualizar',
    'atualizando',
    'atualiza',
    'edite',
    'editar',
    'editando',
    'edita',
    'altere',
    'alterar',
    'alterando',
    'altera',
    'mude',
    'mudar',
    'mudando',
    'muda',
    'modifique',
    'modificar',
    'modificando',
    'modifica',
    'corrija',
    'corrigir',
    'corrige',
    'atualize o lead',
    'atualiza o lead',
    'update lead',
    'update no lead',
    'mudar dado',
    'mudar dados',
    'corrigir dado',
    'corrigir dados',
    'trocar dado',
    'trocar dados',
    'troca dado',
    'troca dados',
    'muda dado',
    'muda dados',
    'altera dado',
    'altera dados',
    'alterar dado',
    'alterar dados',
    'novo telefone',
    'novo email',
    'novo nome',
    'novo fone',
    'trocar telefone',
    'trocar email',
    'trocar nome',
    'trocar fone',
    'change lead',
    'patch lead',
    'atualizar cadastro',
    'editar cadastro',
    'alterar cadastro',
    'modificar cadastro',
    'mudar cadastro',
    'corrigir cadastro',
    'refazer cadastro',
    'atualizar ficha',
    'editar ficha',
    'alterar ficha',
    'fix lead',
    'consertar lead',
    'conserta lead',
    'corrige lead',
    'ajustar lead',
    'ajusta lead',
    'ajuste lead',
  ],

  DELETAR: [
    'exclua',
    'excluir',
    'excluindo',
    'exclui',
    'delete',
    'deletar',
    'deletando',
    'deleta',
    'remova',
    'remover',
    'removendo',
    'remove',
    'apague',
    'apagar',
    'apagando',
    'apaga',
    'limpar lead',
    'limpa lead',
    'limpe lead',
    'matar lead',
    'mata lead',
    'mata o lead',
    'sumir com',
    'sumiu o lead',
    'eliminar lead',
    'elimina lead',
    'elimine lead',
    'dar baixa',
    'dá baixa',
    'da baixa no lead',
    'tirar lead',
    'tira lead',
    'tire lead',
    'desativar lead',
    'desativa lead',
    'desative lead',
    'inativar lead',
    'inativa lead',
    'inative lead',
    'descartar lead',
    'descarta lead',
    'descarte lead',
    'jogar fora',
    'joga fora',
    'jogue fora',
    'remover cadastro',
    'excluir cadastro',
    'deletar cadastro',
    'apagar cadastro',
    'zerar lead',
    'zera lead',
    'zerando lead',
    'drop lead',
    'remove lead',
    'exclusão',
    'exclusao',
    'deletar ficha',
    'apagar ficha',
    'excluir ficha',
  ],

  INTERACAO: [
    'registre',
    'registrar',
    'registrando',
    'registra',
    'anote',
    'anotar',
    'anotando',
    'anota',
    'agende',
    'agendar',
    'agendando',
    'agenda',
    'observação',
    'observacao',
    'observa',
    'obs',
    'nota',
    'histórico',
    'historico',
    'comentário',
    'comentario',
    'comente',
    'comenta',
    'comentar',
    'botar nota',
    'bota nota',
    'coloca nota',
    'colocar nota',
    'inserir nota',
    'insere nota',
    'insira nota',
    'adicionar nota',
    'adiciona nota',
    'adicione nota',
    'registrar que',
    'registre que',
    'anotar que',
    'anote que',
    'lembrete',
    'lembra',
    'lembrar',
    'memorandum',
    'memorando',
    'retorno',
    'agendar retorno',
    'agenda retorno',
    'agendando retorno',
    'agendar visita',
    'agenda visita',
    'callback',
    'follow up',
    'followup',
    'follow-up',
    'marcar',
    'marca',
    'marque',
    'agendar ligação',
    'agenda ligação',
    'anota aí',
    'anota ai',
    'botar no histórico',
    'bota no historico',
    'add nota',
    'add comentario',
    'add observacao',
    'add interação',
    'interação',
    'interacao',
    'registrar interação',
    'interagir',
    'add histórico',
    'add historico',
    'status',
    'alterar status',
    'mudar status',
    'atualizar status',
    'muda status',
    'set status',
  ],
};

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// MOTOR DE DETECÇÃO DE INTENÇÃO — alta performance com indexOf
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function detectarIntencoes(textoMinusculo) {
  const detectadas = [];
  for (const [nome, gatilhos] of Object.entries(INTENCOES)) {
    if (gatilhos.some((g) => textoMinusculo.includes(g))) {
      detectadas.push(nome);
    }
  }
  return detectadas;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// GOVERNANÇA — Verifica dono do lead (cache local + API como fallback)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async function verificarPropriedadeLead(leadId) {
  // 1. Tentar cache local
  const local = bancoDadosCRM.find((l) => l.id === leadId);
  if (local) {
    const autorizado =
      local.corretorId === CORRETOR_ID ||
      (local.responsavel && local.responsavel.toLowerCase() === CORRETOR_NOME.toLowerCase());
    return { encontrado: true, autorizado, lead: local, fonte: 'CACHE_LOCAL' };
  }

  // 2. Fallback para API — busca a lista completa e filtra por ID
  //    (o mock não tem GET /leads/:id, apenas GET /leads)
  const resultado = await enviarParaCRM('GET', '/leads');
  if (!resultado.sucesso) {
    return { encontrado: false, autorizado: false, lead: null, fonte: 'API' };
  }

  const lista = resultado.dados?.dados || resultado.dados?.leads || resultado.dados;
  if (!Array.isArray(lista)) {
    return { encontrado: false, autorizado: false, lead: null, fonte: 'API' };
  }

  const leadApi = lista.find((l) => l.id === leadId);
  if (!leadApi) {
    return { encontrado: false, autorizado: false, lead: null, fonte: 'API' };
  }

  const autorizado =
    leadApi.idcorretor === CORRETOR_ID ||
    leadApi.id_corretor === CORRETOR_ID ||
    leadApi.corretorId === CORRETOR_ID ||
    (leadApi.corretor_nome &&
      leadApi.corretor_nome.toLowerCase() === CORRETOR_NOME.toLowerCase()) ||
    (leadApi.responsavel && leadApi.responsavel.toLowerCase() === CORRETOR_NOME.toLowerCase()) ||
    (leadApi.nome_corretor && leadApi.nome_corretor.toLowerCase() === CORRETOR_NOME.toLowerCase());

  return { encontrado: true, autorizado, lead: leadApi, fonte: 'API' };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// MENSAGENS VISUAIS PADRONIZADAS — UX Premium v3.1
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const MSG = {
  // ── Existentes (mantidas sem alteração) ───────────────────────────────────
  manutencao: () =>
    `> 🤖 **Assistente em Manutenção**\n> Isolei uma instabilidade no seu comando. Por favor, tente novamente com uma instrução mais objetiva.\n\n*Conformidade: Regra 9 — Responder sempre ao usuário*`,

  colisao: (intencoes) =>
    `> ⚠️ **Comando Ambíguo Detectado — Conformidade Pratti**\n\n` +
    `Identifiquei **${intencoes.length} intenções conflitantes** na mesma mensagem: **${intencoes.join(' + ')}**.\n\n` +
    `Por segurança, o sistema aceita apenas **uma ação por mensagem**.\n\n` +
    `📌 *Envie um comando de cada vez. Exemplo:*\n` +
    `• *"Criar lead Maria Souza, fone 45999881122, email maria@test.com"*\n` +
    `• *"Deletar lead 105"*`,

  semInput: () =>
    `> ⚠️ **Mensagem Vazia**\n\nNão recebi nenhum texto para processar. Por favor, envie um comando.`,

  inputMuitoLongo: () =>
    `> ⚠️ **Entrada Muito Longa**\n\nSeu comando foi truncado para ${LIMITE_CARACTERES} caracteres. Por favor, simplifique a instrução.`,

  bloqueioPrivacidade: (leadId, responsavel) =>
    `> 🔒 **Acesso Bloqueado — Conformidade Pratti**\n\n` +
    `**Lead ID ${leadId}** está atribuído ao corretor **${responsavel || 'outro usuário'}**.\n\n` +
    `Conforme as **Regras 3 e 4 de Governança**, você só pode visualizar e editar leads da sua carteira.\n\n` +
    `---\n*Tentativa registrada em log de auditoria.*`,

  leadNaoEncontrado: (leadId) =>
    `> ❌ **Lead Não Localizado**\n\nNenhum lead com **ID ${leadId}** foi encontrado no CRM ou na sua carteira. Verifique o número e tente novamente.`,

  camposObrigatoriosAusentes: (campos) =>
    `> ❌ **Campos Obrigatórios Ausentes — Regra 1**\n\n` +
    `Para concluir o cadastro, preciso dos seguintes dados na mesma mensagem:\n\n` +
    campos.map((c) => `• **${c}**`).join('\n') +
    `\n\n📌 *Exemplo completo (mínimo):*\n` +
    `*"Cadastrar lead Roberto Souza, fone 45 99999-1122, email roberto@pratti.com.br"*\n\n` +
    `💡 *Campos opcionais que também posso registrar:*\n` +
    `• Empreendimento de interesse *(ex: "interessado no Carmel")*\n` +
    `• Origem do lead *(ex: "veio do Instagram")*\n` +
    `• Renda aproximada *(ex: "renda de 8 mil")*\n` +
    `• CPF *(ex: "CPF 123.456.789-00")*\n` +
    `• Data de retorno *(ex: "retorno amanhã às 14h")*\n` +
    `• Observação inicial *(ex: "obs: cliente viaja muito")*`,

  // ── NOVA: Boas-vindas ──────────────────────────────────────────────────────
  boasVindas: () => {
    const hora = new Date().getHours();
    const periodo = hora < 12 ? 'Bom dia' : hora < 18 ? 'Boa tarde' : 'Boa noite';
    return (
      `> 👋 **${periodo}, ${CORRETOR_NOME}! Seu assistente Pratti está ativo.**\n\n` +
      `Sou o bot do CRM Pratti para Microsoft Teams. Interpreto linguagem natural e executo ações direto no **CV CRM** — sem formulários, sem cliques.\n\n` +
      `---\n\n` +
      `## 📋 O que posso fazer por você\n\n` +
      `**➕ Criar Lead**\n` +
      `Cadastre um novo cliente diretamente pelo chat, com todos os dados relevantes.\n` +
      `> *"Crie lead Maria Souza, fone 45 99999-0000, email maria@test.com, interessada no Villa Bella, veio do Instagram, renda de 12 mil"*\n\n` +
      `**📋 Listar Leads**\n` +
      `Visualize todos os leads da sua carteira em uma listagem rápida.\n` +
      `> *"Ver meus leads"* · *"Puxar minha carteira"* · *"Listar todos os clientes"*\n\n` +
      `**✏️ Atualizar Lead**\n` +
      `Altere telefone, e-mail, empreendimento ou outros dados de um lead existente.\n` +
      `> *"Atualizar lead 465: novo telefone 45 99111-2233"*\n\n` +
      `**📝 Registrar Nota / Histórico**\n` +
      `Adicione anotações, observações ou agende retornos no histórico do lead.\n` +
      `> *"Anotar no lead 465: cliente pediu retorno na sexta às 10h"*\n\n` +
      `**🗑️ Remover Lead**\n` +
      `Exclua um lead da base quando necessário (ação irreversível).\n` +
      `> *"Excluir lead 465"*\n\n` +
      `---\n` +
      `💡 Digite **"ajuda"** para ver todos os comandos  ·  **"help criar"**, **"help listar"**, **"help atualizar"**, **"help histórico"** ou **"help deletar"** para detalhes de cada função.\n\n` +
      `🔒 *Você só visualiza e edita leads da sua carteira — governança Pratti ativa.*`
    );
  },

  // ── NOVA: Menu de ajuda geral (melhorado) ─────────────────────────────────
  ajuda: () =>
    `> 🤖 **Assistente Pratti Comercial — Menu de Comandos**\n\n` +
    `Interprete comandos em linguagem natural. Veja o que posso fazer:\n\n` +
    `**➕ CRIAR LEAD**\n` +
    `• *"Cadastrar lead Ana Lima, fone 45 99887-6655, email ana@pratti.com.br"*\n` +
    `• *"Novo lead Pedro Almada, 45999001122, pedro@test.com, interessado no Carmel, veio do Facebook, renda 10 mil"*\n\n` +
    `**📋 LISTAR LEADS**\n` +
    `• *"Ver todos os meus leads"* / *"Puxar minha carteira"*\n\n` +
    `**✏️ ATUALIZAR LEAD**\n` +
    `• *"Alterar telefone do lead 465 para 45 99111-2233"*\n` +
    `• *"Editar lead 465: empreendimento Carmel, email novo@test.com"*\n\n` +
    `**🗑️ DELETAR LEAD**\n` +
    `• *"Excluir lead 465"* / *"Apagar lead 465"*\n\n` +
    `**📝 REGISTRAR NOTA/HISTÓRICO**\n` +
    `• *"Anotar no lead 465: cliente solicitou retorno amanhã às 14h"*\n` +
    `• *"Registrar que o lead 465 visitou o stand hoje"*\n\n` +
    `---\n` +
    `💡 **Campos que posso capturar ao criar ou atualizar:**\n` +
    `Nome · Telefone · E-mail · Empreendimento · Origem · Renda · CPF · Data de retorno · Observação inicial\n\n` +
    `📖 Digite **"help criar"**, **"help listar"**, **"help atualizar"**, **"help histórico"** ou **"help deletar"** para detalhes completos.\n\n` +
    `🔒 *Somente leads da sua carteira são visíveis e editáveis.*`,

  // ── NOVA: Help detalhado — CRIAR ──────────────────────────────────────────
  helpCriar: () =>
    `> ➕ **Ajuda Detalhada — Criar Lead**\n\n` +
    `Cadastra um novo cliente no CV CRM e o vincula automaticamente à sua carteira.\n\n` +
    `---\n\n` +
    `**📌 Campos obrigatórios** *(sem eles o cadastro é bloqueado)*\n` +
    `• **Nome completo** do cliente\n` +
    `• **Telefone** *(formatos aceitos: 45999991111, 45 99999-1111, (45)99999-1111)*\n` +
    `• **E-mail** *(formato: exemplo@dominio.com)*\n\n` +
    `---\n\n` +
    `**💡 Campos opcionais** *(enriquecem o lead no CRM)*\n` +
    `• **Empreendimento** — *"interessado no Carmel"*, *"interesse no Villa Bella"*\n` +
    `• **Origem** — *"veio do Instagram"*, *"indicação"*, *"captado no Facebook"*, *"via Google"*, *"plantão"*, *"feirão"*, *"evento"*, *"stand"*\n` +
    `• **Renda aproximada** — *"renda de 12 mil"*, *"renda aproximada 8000"*, *"renda R$ 5.500"*\n` +
    `• **CPF** — *"CPF 123.456.789-00"*\n` +
    `• **Observação inicial** — *"obs: cliente prefere contato por WhatsApp"*\n` +
    `• **Data de retorno** — *"retorno amanhã às 14h"*, *"ligar na sexta às 10h"*\n\n` +
    `---\n\n` +
    `**📋 Exemplos de comando**\n\n` +
    `*Mínimo:*\n` +
    `> *"Crie lead João Pereira, fone 45 99888-7766, email joao@test.com"*\n\n` +
    `*Completo:*\n` +
    `> *"Cadastrar lead Maria Oliveira, 45999001122, maria@test.com, interessada no Carmel, veio do Instagram, renda de 10 mil, CPF 111.222.333-44, obs: cliente viaja muito, retorno na sexta às 14h"*\n\n` +
    `---\n` +
    `🔒 *O lead é criado automaticamente vinculado ao seu usuário (${CORRETOR_NOME}).*`,

  // ── NOVA: Help detalhado — LISTAR ─────────────────────────────────────────
  helpListar: () =>
    `> 📋 **Ajuda Detalhada — Listar Leads**\n\n` +
    `Exibe todos os leads da sua carteira com nome, telefone mascarado, empreendimento e ID.\n\n` +
    `---\n\n` +
    `**📋 Exemplos de comando**\n` +
    `• *"Ver meus leads"*\n` +
    `• *"Listar minha carteira"*\n` +
    `• *"Puxar todos os clientes"*\n` +
    `• *"Mostra minha base"*\n` +
    `• *"Consultar leads"*\n` +
    `• *"Trazer carteira"*\n` +
    `• *"Meus contatos"*\n\n` +
    `---\n` +
    `🔒 *Leads de outros corretores não aparecem na listagem — Regra 4 de Governança.*`,

  // ── NOVA: Help detalhado — ATUALIZAR ──────────────────────────────────────
  helpAtualizar: () =>
    `> ✏️ **Ajuda Detalhada — Atualizar Lead**\n\n` +
    `Altera dados cadastrais de um lead existente da sua carteira.\n\n` +
    `---\n\n` +
    `**📌 Como identificar o lead**\n` +
    `• Pelo **ID numérico** *(mais preciso)*: *"lead 465"*\n` +
    `• Pelo **nome** *(pode gerar desambiguação se houver duplicatas)*: *"lead Maria Souza"*\n\n` +
    `**💡 O que posso atualizar**\n` +
    `• **Nome** — *"nome para Maria Souza Lima"*\n` +
    `• **Telefone** — *"fone 45 99111-2233"*\n` +
    `• **E-mail** — *"email novo@test.com"*\n` +
    `• **Empreendimento** — *"empreendimento Carmel"*\n` +
    `• **Origem** — *"origem Instagram"*\n` +
    `• **Renda** — *"renda de 15 mil"*\n\n` +
    `---\n\n` +
    `**📋 Exemplos de comando**\n` +
    `• *"Atualizar lead 465: telefone 45 99111-2233"*\n` +
    `• *"Editar lead Maria Souza: novo email maria2@test.com, empreendimento Villa Bella"*\n` +
    `• *"Alterar lead 465: renda de 18 mil, origem Google"*\n\n` +
    `---\n` +
    `🔒 *Somente leads da sua carteira podem ser editados — Regra 3 de Governança.*`,

  // ── NOVA: Help detalhado — HISTÓRICO/INTERAÇÃO ────────────────────────────
  helpHistorico: () =>
    `> 📝 **Ajuda Detalhada — Registrar Nota / Histórico**\n\n` +
    `Adiciona uma anotação ou observação no histórico de atendimento do lead, com registro de autor, data e hora automaticamente.\n\n` +
    `---\n\n` +
    `**📌 Como identificar o lead**\n` +
    `• Pelo **ID numérico** *(mais preciso)*: *"lead 465"*\n` +
    `• Pelo **nome** *(pode gerar desambiguação se houver duplicatas)*\n\n` +
    `**💡 O que posso registrar**\n` +
    `• Observações livres sobre o atendimento\n` +
    `• Resultado de ligações e visitas\n` +
    `• Agendamento de retorno\n` +
    `• Status do interesse do cliente\n\n` +
    `---\n\n` +
    `**📋 Exemplos de comando**\n` +
    `• *"Anotar no lead 465: cliente pediu retorno na sexta às 10h"*\n` +
    `• *"Registrar no lead 465 que o cliente visitou o stand hoje e gostou do Carmel"*\n` +
    `• *"Obs no lead 465: sem resposta na ligação, tentar novamente amanhã"*\n` +
    `• *"Histórico lead 465: cliente aprovado no financiamento, aguardando documentação"*\n\n` +
    `---\n` +
    `🔒 *A nota é registrada com seu nome (${CORRETOR_NOME}), origem Teams e timestamp automático — Regra 7.*`,

  // ── NOVA: Help detalhado — DELETAR ────────────────────────────────────────
  helpDeletar: () =>
    `> 🗑️ **Ajuda Detalhada — Deletar Lead**\n\n` +
    `Remove permanentemente um lead do CRM. **Esta ação é irreversível.**\n\n` +
    `---\n\n` +
    `**📌 Como identificar o lead**\n` +
    `• Recomendamos sempre usar o **ID numérico** para evitar erros:\n` +
    `  *"Excluir lead 465"*\n\n` +
    `**📋 Exemplos de comando**\n` +
    `• *"Excluir lead 465"*\n` +
    `• *"Deletar lead 465"*\n` +
    `• *"Apagar lead 465"*\n` +
    `• *"Remover lead 465"*\n\n` +
    `---\n` +
    `⚠️ *Se houver dúvida entre dois leads com o mesmo nome, o bot pedirá confirmação antes de excluir.*\n` +
    `🔒 *Somente leads da sua carteira podem ser removidos — Regra 3 de Governança.*`,
};

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
//  HANDLERS DE INTENÇÃO — cada um em seu próprio try/catch
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

async function handleCriar(textoOriginal, textoMinusculo, send) {
  try {
    const telefone = extrairTelefone(textoOriginal);
    const email = extrairEmail(textoOriginal);
    const nome = extrairNome(textoOriginal);

    const camposFaltando = [];
    if (!nome) camposFaltando.push('Nome do cliente');
    if (!telefone) camposFaltando.push('Telefone');
    if (!email) camposFaltando.push('E-mail');

    if (camposFaltando.length > 0) {
      registrarLog(
        'BLOQUEIO',
        CORRETOR_NOME,
        'CRIAR_LEAD_CAMPOS_AUSENTES',
        { camposFaltando },
        `Ação bloqueada: campos obrigatórios ausentes — ${camposFaltando.join(', ')}.`
      );
      await send(MSG.camposObrigatoriosAusentes(camposFaltando));
      return;
    }

    // Extrai empreendimento (opcional)
    const matchEmp = textoOriginal.match(
      /(?:empreendimento|empre?end|interesse\s+em|interessad[ao]\s+(?:no|na|em))\s+([A-ZÀ-Ÿa-zà-ÿ ]{3,30})/i
    );
    const empreendimento = matchEmp ? matchEmp[1].trim() : 'Não informado';

    // Extrai corretor mencionado (se houver)
    const matchCorretor = textoOriginal.match(
      /corretor\s+([A-ZÀ-Ÿa-zà-ÿ]+(?:\s+[A-ZÀ-Ÿa-zà-ÿ]+)?)/i
    );
    const corretor = matchCorretor ? matchCorretor[1].trim() : CORRETOR_NOME;

    const origem = extrairOrigem(textoOriginal);
    const renda = extrairRenda(textoOriginal);
    const cpf = extrairCPF(textoOriginal);
    const dataRetorno = extrairDataRetorno(textoOriginal);
    const obsInicial = extrairObservacaoInicial(textoOriginal);

    const telefoneLimpo = telefone.replace(/\D/g, '');
    const emailLimpo = email.toLowerCase();

    await send(`⏳ Verificando duplicidade no CRM para **${nome}**...`);

    // Verificação de duplicidade
    const checagem = await enviarParaCRM('GET', '/leads');
    if (checagem.sucesso) {
      const lista = checagem.dados?.dados || checagem.dados?.leads || checagem.dados;
      if (Array.isArray(lista)) {
        const duplicado = lista.find((l) => {
          const lEmail = (l.email || '').trim().toLowerCase();
          const lFone = (l.telefone || '').replace(/\D/g, '');
          return lEmail === emailLimpo || lFone === telefoneLimpo;
        });
        if (duplicado) {
          registrarLog(
            'BLOQUEIO',
            CORRETOR_NOME,
            'CRIAR_LEAD_DUPLICADO',
            { id: duplicado.id, email: emailLimpo },
            `Ação bloqueada: lead duplicado localizado com ID ${duplicado.id}.`
          );
          await send(
            `> ⚠️ **Cadastro Duplicado Detectado — Regra 1**\n\n` +
              `O cliente **${nome}** já possui cadastro ativo no CRM.\n\n` +
              `• **ID Existente:** ${duplicado.id}\n` +
              `• **Nome:** ${duplicado.nome || 'Não informado'}\n` +
              `• **E-mail:** ${duplicado.email || 'Não informado'}\n` +
              `• **Responsável:** ${duplicado.responsavel || 'Fila Geral'}\n\n` +
              `*Para atualizar, use: "Atualizar lead ${duplicado.id} ..."*`
          );
          return;
        }
      }
    }

    // Dentro do objeto payload existente, adicione as novas propriedades:
    const payload = {
      nome,
      telefone,
      email: emailLimpo,
      empreendimento,
      origem: origem || 'Teams Bot Pratti', // ← campo já existia, agora usa o extrator
      renda_aproximada: renda || null, // ← NOVO
      cpf: cpf || null, // ← NOVO
      data_retorno: dataRetorno || null, // ← NOVO
      observacao_inicial: obsInicial || null, // ← NOVO
      gestor_responsavel: CORRETOR_NOME,
      id_corretor: CORRETOR_ID,
      corretor,
      repassar: 'N',
      imobiliaria: 'Pratti',
    };

    await send(`📡 Enviando cadastro para o servidor do CV CRM...`);

    const resultado = await enviarParaCRM('POST', '/leads', payload);

    if (resultado.sucesso) {
      const novoId =
        resultado.dados?.id ||
        resultado.dados?.idlead ||
        Math.max(...bancoDadosCRM.map((l) => l.id)) + 1;
      bancoDadosCRM.push({
        ...payload,
        id: novoId,
        responsavel: CORRETOR_NOME,
        corretorId: CORRETOR_ID,
      });
      registrarLog(
        'SUCESSO',
        CORRETOR_NOME,
        'CRIAR_LEAD_OK',
        { ...payload, novoId },
        `Lead criado e atribuído ao corretor ${CORRETOR_NOME} (ID ${CORRETOR_ID}).`
      );
      // Substituir o await send de sucesso existente por este:
      const linhasOpcionais = [
        empreendimento !== 'Não informado' ? `• **Empreendimento:** ${empreendimento}` : null,
        origem ? `• **Origem:** ${origem}` : null,
        renda ? `• **Renda Aprox.:** ${renda}` : null,
        cpf ? `• **CPF:** ${cpf}` : null,
        dataRetorno ? `• **Retorno Agendado:** ${dataRetorno}` : null,
        obsInicial ? `• **Obs. Inicial:** ${obsInicial}` : null,
      ].filter(Boolean);

      await send(
        `> ✅ **Lead Criado com Sucesso — CV CRM**\n\n` +
          `**Dados obrigatórios:**\n` +
          `• **Nome:** ${nome}\n` +
          `• **Telefone:** ${telefone}\n` +
          `• **E-mail:** ${emailLimpo}\n` +
          `• **ID Gerado:** ${novoId}\n` +
          `• **Responsável:** ${CORRETOR_NOME} (ID ${CORRETOR_ID})\n\n` +
          (linhasOpcionais.length > 0
            ? `**Dados complementares registrados:**\n` + linhasOpcionais.join('\n') + `\n\n`
            : `💡 *Dica: você pode incluir empreendimento, origem, renda, CPF, data de retorno e observação no mesmo comando.*\n\n`) +
          `---\n*Lead vinculado à sua carteira conforme Regra 2.*`
      );
    } else {
      registrarLog(
        'ERRO',
        CORRETOR_NOME,
        'CRIAR_LEAD_FALHA_API',
        { erro: resultado.erro },
        `API recusou a criação do lead.`
      );
      await send(`> ❌ **CRM Recusou o Cadastro**\n\n${resultado.erro}`);
    }
  } catch (err) {
    registrarLog(
      'FATAL',
      CORRETOR_NOME,
      'CRIAR_LEAD_EXCEPTION',
      { stack: err.stack },
      'Exceção interna capturada no handler CRIAR.'
    );
    await send(MSG.manutencao());
  }
}

// ─────────────────────────────────────────────────────────────────────────────

async function handleListar(send) {
  try {
    await send(`⏳ Consultando sua carteira de leads no CRM...`);

    const resultado = await enviarParaCRM('GET', '/leads');

    if (!resultado.sucesso) {
      await send(`> ❌ **Falha na Consulta**\n\n${resultado.erro}`);
      return;
    }

    const lista = resultado.dados?.dados || resultado.dados?.leads || resultado.dados;

    if (!Array.isArray(lista)) {
      // Fallback para banco local
      const minha = bancoDadosCRM.filter(
        (l) =>
          l.corretorId === CORRETOR_ID ||
          (l.responsavel || '').toLowerCase() === CORRETOR_NOME.toLowerCase()
      );
      if (minha.length === 0) {
        await send(`> 📭 **Carteira Vazia**\n\nNenhum lead ativo encontrado na sua carteira.`);
        return;
      }
      await send(montarMensagemLista(minha, true));
      return;
    }

    // DIRETIVA 5 — FILTRO DE PRIVACIDADE: exibe SOMENTE leads do corretor logado
    // Cobre todos os formatos de campo: mock API usa idcorretor + corretor_nome
    const minha = lista.filter(
      (l) =>
        l.idcorretor === CORRETOR_ID ||
        l.id_corretor === CORRETOR_ID ||
        l.corretorId === CORRETOR_ID ||
        (l.corretor_nome && l.corretor_nome.toLowerCase() === CORRETOR_NOME.toLowerCase()) ||
        (l.responsavel && l.responsavel.toLowerCase() === CORRETOR_NOME.toLowerCase()) ||
        (l.nome_corretor && l.nome_corretor.toLowerCase() === CORRETOR_NOME.toLowerCase())
    );

    registrarLog(
      'INFO',
      CORRETOR_NOME,
      'LISTAR_LEADS',
      { total_api: lista.length, total_visivel: minha.length },
      `Filtro de privacidade aplicado. Exibindo apenas leads do corretor ${CORRETOR_NOME} (ID ${CORRETOR_ID}).`
    );

    if (minha.length === 0) {
      await send(`> 📭 **Carteira Vazia**\n\nNenhum lead atribuído à sua carteira no momento.`);
      return;
    }

    await send(montarMensagemLista(minha, false));
  } catch (err) {
    registrarLog(
      'FATAL',
      CORRETOR_NOME,
      'LISTAR_LEADS_EXCEPTION',
      { stack: err.stack },
      'Exceção interna capturada no handler LISTAR.'
    );
    await send(MSG.manutencao());
  }
}

function montarMensagemLista(leads, isLocal) {
  const fonte = isLocal ? ' *(dados do cache local)*' : '';
  let msg = `> 📋 **Sua Carteira de Leads — CV CRM**${fonte}\n\n`;
  const exibir = leads.slice(0, 20);
  exibir.forEach((l, i) => {
    const nome = l.nome || l.nome_cliente || 'Sem Nome';
    const fone = l.telefone || l.celular || 'Não informado';
    const emp = l.empreendimento || l.empreendimento_interesse || 'Não informado';
    msg += `**${i + 1}. ID ${l.id}** — ${nome}\n• Fone: ${fone}  |  Empreendimento: ${emp}\n\n`;
  });
  if (leads.length > 20)
    msg += `---\n*Exibindo 20 de ${leads.length} leads. Use o ID para ações específicas.*`;
  return msg;
}

// ─────────────────────────────────────────────────────────────────────────────

async function handleAtualizar(textoOriginal, userId, send) {
  try {
    const ids = extrairIds(textoOriginal);

    if (!ids || ids.length === 0) {
      // Tenta desambiguação por nome
      const nome = extrairNome(textoOriginal);
      if (nome) {
        const candidatos = bancoDadosCRM.filter(
          (l) =>
            (l.corretorId === CORRETOR_ID ||
              (l.responsavel || '').toLowerCase() === CORRETOR_NOME.toLowerCase()) &&
            l.nome.toLowerCase().includes(nome.toLowerCase())
        );
        if (candidatos.length === 0) {
          await send(
            `> ❌ **Lead Não Encontrado**\n\nNenhum lead com o nome **"${nome}"** foi localizado na sua carteira.`
          );
          return;
        }
        if (candidatos.length === 1) {
          // Confirma automaticamente e segue
          return await executarAtualizacao(candidatos[0].id, textoOriginal, userId, send);
        }
        // Múltiplos — desambigua
        return await iniciarDesambiguacao(candidatos, userId, 'ATUALIZAR', textoOriginal, send);
      }
      await send(
        `> ❌ **ID Não Identificado**\n\nInforme o **ID numérico** do lead ou o **nome completo**.\n\n` +
          `*Exemplo: "Alterar email do lead 465 para novo@email.com"*`
      );
      return;
    }

    if (ids.length > 1) {
      await send(
        `> ⚠️ **Múltiplos IDs Detectados**\n\nIdentifiquei os IDs: **${ids.join(', ')}**. Informe apenas um ID por comando.`
      );
      return;
    }

    await executarAtualizacao(ids[0], textoOriginal, userId, send);
  } catch (err) {
    registrarLog(
      'FATAL',
      CORRETOR_NOME,
      'ATUALIZAR_LEAD_EXCEPTION',
      { stack: err.stack },
      'Exceção interna capturada no handler ATUALIZAR.'
    );
    await send(MSG.manutencao());
  }
}

async function executarAtualizacao(leadId, textoOriginal, userId, send) {
  const propriedade = await verificarPropriedadeLead(leadId);

  if (!propriedade.encontrado) {
    await send(MSG.leadNaoEncontrado(leadId));
    return;
  }
  if (!propriedade.autorizado) {
    registrarLog(
      'VIOLAÇÃO',
      CORRETOR_NOME,
      'TENTATIVA_EDITAR_LEAD_ALHEIO',
      { leadId, responsavel: propriedade.lead?.responsavel },
      `Ação bloqueada pois o lead ${leadId} pertence ao corretor "${propriedade.lead?.responsavel}" — não ao usuário ${CORRETOR_NOME} (ID ${CORRETOR_ID}).`
    );
    console.log(
      `🚨 [ALERTA DE VIOLAÇÃO] Usuário ${CORRETOR_NOME} tentou editar o lead ${leadId} que pertence a "${propriedade.lead?.responsavel}".`
    );
    await send(MSG.bloqueioPrivacidade(leadId, propriedade.lead?.responsavel));
    return;
  }

  const telefone = extrairTelefone(textoOriginal);
  const email = extrairEmail(textoOriginal);
  const matchNome = textoOriginal.match(
    /\bnome\s+(?:para\s+)?([A-ZÀ-Ÿa-zà-ÿ]+(?:\s+[A-ZÀ-Ÿa-zà-ÿ]+){1,3})/i
  );
  const novoNome = matchNome ? matchNome[1].trim() : null;
  const matchEmp = textoOriginal.match(
    /(?:empreendimento|empre?end)\s+(?:para\s+)?([A-ZÀ-Ÿa-zà-ÿ ]{3,30})/i
  );
  const novoEmp = matchEmp ? matchEmp[1].trim() : null;

  const payloadUpdate = {};
  if (telefone) payloadUpdate.telefone = telefone;
  if (email) payloadUpdate.email = email.toLowerCase();
  if (novoNome) payloadUpdate.nome = novoNome;
  if (novoEmp) payloadUpdate.empreendimento = novoEmp;

  if (Object.keys(payloadUpdate).length === 0) {
    await send(
      `> ⚠️ **Nenhum Campo Novo Identificado**\n\n` +
        `Não encontrei dados para atualizar. Informe o que deseja alterar.\n\n` +
        `*Exemplo: "Alterar email do lead ${leadId} para novo@email.com"*`
    );
    return;
  }

  await send(`⏳ Enviando atualização do lead **ID ${leadId}** para o CRM...`);

  const resultado = await enviarParaCRM('PATCH', `/leads/${leadId}`, payloadUpdate);

  if (resultado.sucesso) {
    // Sincroniza cache local
    const idx = bancoDadosCRM.findIndex((l) => l.id === leadId);
    if (idx !== -1) Object.assign(bancoDadosCRM[idx], payloadUpdate);

    registrarLog(
      'SUCESSO',
      CORRETOR_NOME,
      'ATUALIZAR_LEAD_OK',
      { leadId, ...payloadUpdate },
      `Lead ${leadId} atualizado com sucesso pelo corretor ${CORRETOR_NOME}.`
    );

    const linhas = Object.entries(payloadUpdate)
      .map(([k, v]) => `• **${k}:** ${v}`)
      .join('\n');
    await send(
      `> ✅ **Lead Atualizado com Sucesso**\n\n` +
        `**ID ${leadId}** — Campos alterados:\n\n${linhas}\n\n` +
        `---\n*Registrado em ${new Date().toLocaleString('pt-BR')}*`
    );
  } else {
    registrarLog(
      'ERRO',
      CORRETOR_NOME,
      'ATUALIZAR_LEAD_FALHA_API',
      { leadId, erro: resultado.erro },
      'API recusou a atualização.'
    );
    await send(`> ❌ **CRM Recusou a Atualização**\n\n${resultado.erro}`);
  }
}

// ─────────────────────────────────────────────────────────────────────────────

async function handleDeletar(textoOriginal, userId, send) {
  try {
    const ids = extrairIds(textoOriginal);

    if (!ids || ids.length === 0) {
      await send(
        `> ❌ **ID Não Identificado**\n\nInforme o **ID numérico** do lead para exclusão.\n\n` +
          `*Exemplo: "Deletar lead 465"*`
      );
      return;
    }
    if (ids.length > 1) {
      await send(
        `> ⚠️ **Múltiplos IDs Detectados**\n\nIdentifiquei os IDs: **${ids.join(', ')}**. Por segurança, informe apenas um ID por exclusão.`
      );
      return;
    }

    const leadId = ids[0];
    const propriedade = await verificarPropriedadeLead(leadId);

    if (!propriedade.encontrado) {
      await send(MSG.leadNaoEncontrado(leadId));
      return;
    }
    if (!propriedade.autorizado) {
      registrarLog(
        'VIOLAÇÃO',
        CORRETOR_NOME,
        'TENTATIVA_DELETAR_LEAD_ALHEIO',
        { leadId, responsavel: propriedade.lead?.responsavel },
        `Ação bloqueada pois o lead ${leadId} pertence ao corretor "${propriedade.lead?.responsavel}".`
      );
      console.log(
        `🚨 [ALERTA DE VIOLAÇÃO] Usuário ${CORRETOR_NOME} tentou DELETAR o lead ${leadId} que pertence a "${propriedade.lead?.responsavel}".`
      );
      await send(MSG.bloqueioPrivacidade(leadId, propriedade.lead?.responsavel));
      return;
    }

    await send(`⏳ Solicitando exclusão do lead **ID ${leadId}** ao CRM...`);

    const resultado = await enviarParaCRM('DELETE', `/leads/${leadId}`);

    if (resultado.sucesso) {
      bancoDadosCRM = bancoDadosCRM.filter((l) => l.id !== leadId);
      registrarLog(
        'SUCESSO',
        CORRETOR_NOME,
        'DELETAR_LEAD_OK',
        { leadId },
        `Lead ${leadId} excluído com sucesso pelo corretor ${CORRETOR_NOME}.`
      );
      await send(
        `> 🗑️ **Lead Excluído com Sucesso**\n\n` +
          `O lead **ID ${leadId}** foi removido do painel do CV CRM.\n\n` +
          `---\n*Operação registrada em ${new Date().toLocaleString('pt-BR')}*`
      );
    } else {
      registrarLog(
        'ERRO',
        CORRETOR_NOME,
        'DELETAR_LEAD_FALHA_API',
        { leadId, erro: resultado.erro },
        'API recusou a exclusão.'
      );
      await send(`> ❌ **CRM Barrou a Exclusão**\n\n${resultado.erro}`);
    }
  } catch (err) {
    registrarLog(
      'FATAL',
      CORRETOR_NOME,
      'DELETAR_LEAD_EXCEPTION',
      { stack: err.stack },
      'Exceção interna capturada no handler DELETAR.'
    );
    await send(MSG.manutencao());
  }
}

// ─────────────────────────────────────────────────────────────────────────────

async function handleInteracao(textoOriginal, userId, send) {
  try {
    const ids = extrairIds(textoOriginal);

    if (!ids || ids.length === 0) {
      await send(
        `> ❌ **ID Não Identificado**\n\nInforme o **ID do lead** para registrar a nota.\n\n` +
          `*Exemplo: "Anotar no lead 465: cliente solicitou retorno às 14h"*`
      );
      return;
    }

    const leadId = ids[0];
    const propriedade = await verificarPropriedadeLead(leadId);

    if (!propriedade.encontrado) {
      await send(MSG.leadNaoEncontrado(leadId));
      return;
    }
    if (!propriedade.autorizado) {
      registrarLog(
        'VIOLAÇÃO',
        CORRETOR_NOME,
        'TENTATIVA_INTERACAO_LEAD_ALHEIO',
        { leadId, responsavel: propriedade.lead?.responsavel },
        `Ação bloqueada pois o lead ${leadId} pertence ao corretor "${propriedade.lead?.responsavel}".`
      );
      console.log(
        `🚨 [ALERTA DE VIOLAÇÃO] Usuário ${CORRETOR_NOME} tentou registrar interação no lead ${leadId} que pertence a "${propriedade.lead?.responsavel}".`
      );
      await send(MSG.bloqueioPrivacidade(leadId, propriedade.lead?.responsavel));
      return;
    }

    // Extrai o texto da observação removendo os gatilhos e o ID
    let descricao = textoOriginal;
    INTENCOES.INTERACAO.forEach((g) => {
      descricao = descricao.replace(
        new RegExp(`\\b${g.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'gi'),
        ' '
      );
    });
    descricao = descricao
      .replace(/\b(?:no|do|para\s+o|no\s+lead|id)\s*\d+/gi, ' ')
      .replace(/\b\d{2,6}\b/g, ' ')
      .replace(/[-_:;,]+/g, ' ')
      .replace(/\s{2,}/g, ' ')
      .trim();

    if (!descricao || descricao.length < 3) {
      descricao = 'Acompanhamento registrado via Microsoft Teams.';
    }

    // Detecta agendamento de data/hora
    const matchData = textoOriginal.match(/\b(\d{1,2}[\/\-]\d{1,2}(?:[\/\-]\d{2,4})?)\b/);
    const matchHora = textoOriginal.match(/\b(\d{1,2}[h:]\d{0,2})\b/i);
    const infoData = matchData ? ` | Data mencionada: ${matchData[1]}` : '';
    const infoHora = matchHora ? ` | Hora mencionada: ${matchHora[1]}` : '';

    const payload = {
      idlead: leadId,
      usuario: CORRETOR_NOME,
      id_corretor: CORRETOR_ID,
      origem: 'Microsoft Teams — Bot Pratti',
      data_hora: new Date().toISOString(),
      descricao: descricao + infoData + infoHora,
    };

    registrarLog(
      'INFO',
      CORRETOR_NOME,
      'REGISTRAR_INTERACAO',
      payload,
      `Payload de interação montado para o lead ${leadId}.`
    );
    await send(`⏳ Gravando histórico no lead **ID ${leadId}**...`);

    const resultado = await enviarParaCRM('PATCH', `/leads/${leadId}`, payload);

    if (resultado.sucesso) {
      registrarLog(
        'SUCESSO',
        CORRETOR_NOME,
        'INTERACAO_OK',
        { leadId, descricao: payload.descricao },
        `Histórico gravado com sucesso no lead ${leadId}.`
      );
      await send(
        `> 📝 **Histórico Gravado com Sucesso**\n\n` +
          `• **Lead ID:** ${leadId}\n` +
          `• **Usuário:** ${CORRETOR_NOME}\n` +
          `• **Origem:** Microsoft Teams\n` +
          `• **Anotação:** *"${payload.descricao}"*\n` +
          `• **Data/Hora:** ${new Date().toLocaleString('pt-BR')}\n\n` +
          `---\n*Conforme Regra 7 — Observações registradas com origem, usuário e timestamp.*`
      );
    } else {
      registrarLog(
        'ERRO',
        CORRETOR_NOME,
        'INTERACAO_FALHA_API',
        { leadId, erro: resultado.erro },
        'API recusou o registro de interação.'
      );
      await send(`> ❌ **Falha ao Gravar no CRM**\n\n${resultado.erro}`);
    }
  } catch (err) {
    registrarLog(
      'FATAL',
      CORRETOR_NOME,
      'INTERACAO_EXCEPTION',
      { stack: err.stack },
      'Exceção interna capturada no handler INTERACAO.'
    );
    await send(MSG.manutencao());
  }
}

// ─────────────────────────────────────────────────────────────────────────────

async function iniciarDesambiguacao(candidatos, userId, acaoPendente, textoPendente, send) {
  try {
    const opcoes = candidatos.slice(0, 5);
    definirEstado(userId, {
      etapa: 'AGUARDANDO_ID',
      candidatos: opcoes,
      acaoPendente,
      textoPendente,
    });

    let msg =
      `> ⚠️ **Múltiplos Leads Encontrados — Regra 5**\n\n` +
      `Encontrei **${opcoes.length} leads** que correspondem à sua busca. Qual deles deseja editar?\n\n`;

    opcoes.forEach((l, i) => {
      const fone = l.telefone ? `****${String(l.telefone).slice(-4)}` : 'Não informado';
      const emp = l.empreendimento || 'Não informado';
      msg += `**${i + 1}.** ID ${l.id} — ${l.nome}\n• Fone: ${fone}  |  Empreendimento: ${emp}\n\n`;
    });

    msg += `---\n*Digite o número da opção (1 a ${opcoes.length}) ou "cancelar".*`;
    await send(msg);
  } catch (err) {
    registrarLog(
      'FATAL',
      CORRETOR_NOME,
      'DESAMBIGUACAO_EXCEPTION',
      { stack: err.stack },
      'Exceção no handler de desambiguação.'
    );
    await send(MSG.manutencao());
  }
}

async function handleRespostaDesambiguacao(textoOriginal, estado, userId, send) {
  try {
    if (textoOriginal.toLowerCase() === 'cancelar') {
      limparEstado(userId);
      await send(
        `> ℹ️ **Operação Cancelada**\n\nNenhuma alteração foi realizada. Como posso ajudar?`
      );
      return;
    }

    const num = parseInt(textoOriginal.trim(), 10);
    if (isNaN(num) || num < 1 || num > estado.candidatos.length) {
      await send(
        `> ⚠️ **Opção Inválida**\n\n` +
          `Digite um número de **1 a ${estado.candidatos.length}** ou **"cancelar"** para abortar.`
      );
      return;
    }

    const escolhido = estado.candidatos[num - 1];
    limparEstado(userId);

    registrarLog(
      'INFO',
      CORRETOR_NOME,
      'DESAMBIGUACAO_RESOLVIDA',
      { leadId: escolhido.id },
      `Usuário escolheu o lead ${escolhido.id} após desambiguação.`
    );

    // Executa a ação pendente com o ID resolvido
    if (estado.acaoPendente === 'ATUALIZAR') {
      await executarAtualizacao(escolhido.id, estado.textoPendente, userId, send);
    } else if (estado.acaoPendente === 'INTERACAO') {
      await handleInteracao(`${estado.textoPendente} ${escolhido.id}`, userId, send);
    } else if (estado.acaoPendente === 'DELETAR') {
      await handleDeletar(`deletar lead ${escolhido.id}`, userId, send);
    }
  } catch (err) {
    registrarLog(
      'FATAL',
      CORRETOR_NOME,
      'RESPOSTA_DESAMBIGUACAO_EXCEPTION',
      { stack: err.stack },
      'Exceção ao processar resposta de desambiguação.'
    );
    await send(MSG.manutencao());
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// NÚCLEO PRINCIPAL — app.on('message')
// Toda a lógica de roteamento passa por aqui, blindada em try/catch total.
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
app.on('message', async ({ send, activity }) => {
  try {
    const userId = activity?.from?.id || 'UNKNOWN_USER';

    // ── DIRETIVA 4: Saneamento e validação de input ──
    const textoRaw = activity?.text ?? '';
    const textoOriginal = sanitizarInput(textoRaw);

    if (!textoOriginal || textoOriginal.length === 0) {
      await send(MSG.semInput());
      return;
    }

    // Avisa se truncou
    if (textoRaw.length > LIMITE_CARACTERES) {
      registrarLog(
        'WARN',
        CORRETOR_NOME,
        'INPUT_TRUNCADO',
        { tamanhoOriginal: textoRaw.length },
        `Mensagem truncada de ${textoRaw.length} para ${LIMITE_CARACTERES} caracteres.`
      );
      await send(MSG.inputMuitoLongo());
    }

    const textoMinusculo = textoOriginal.toLowerCase();

    registrarLog(
      'INFO',
      CORRETOR_NOME,
      'MENSAGEM_RECEBIDA',
      { texto: textoOriginal },
      `Nova mensagem processada do usuário ${userId}.`
    );

    // ── DIRETIVA 7: Verifica estado de conversa ativo ──
    const estado = obterEstado(userId);
    if (estado && estado.etapa === 'AGUARDANDO_ID') {
      await handleRespostaDesambiguacao(textoOriginal, estado, userId, send);
      return;
    }

    // ── DIRETIVA 2 + 4: Detecção de intenções e colisão ──
    const intencoesDetectadas = detectarIntencoes(textoMinusculo);

    // DEPOIS (substituir pelo bloco abaixo):

    // ── Saudações puras → Boas-vindas ──
    if (
      textoMinusculo.match(
        /^(oi|olá|ola|oi bot|olá bot|ola bot|hey|ei|eai|eaí|e aí|e ai|bom dia|boa tarde|boa noite|good morning|good afternoon|good evening|hello|hi|howdy|salve|salve bot|fala bot|fala aí|fala ai|opa|opa bot|oi tudo bem|tudo bem|tudo bom|tudo certo|tudo ótimo|tudo otimo|como vai|como você está|como voce esta|olá tudo|oi tudo|boas|cumprimentos|saudações|saudacoes|alô|alo|heyy|hola|yoo|yooo|bom diaa|boa tardee|boa noitee|oi sumido|sumido|voltei|to aqui|tô aqui|presente|opa tudo|ola tudo|salutations)\b/
      )
    ) {
      await send(MSG.boasVindas());
      return;
    }

    // ── Help detalhado por função ──
    if (textoMinusculo.match(/\b(help|ajuda|socorro)\b.*\b(criar?|cadastr|novo lead|add lead)\b/)) {
      await send(MSG.helpCriar());
      return;
    }
    if (
      textoMinusculo.match(
        /\b(help|ajuda|socorro)\b.*\b(listar?|list|ver|mostrar?|consultar?|carteira)\b/
      )
    ) {
      await send(MSG.helpListar());
      return;
    }
    if (
      textoMinusculo.match(
        /\b(help|ajuda|socorro)\b.*\b(atualizar?|editar?|alterar?|modificar?|update)\b/
      )
    ) {
      await send(MSG.helpAtualizar());
      return;
    }
    if (
      textoMinusculo.match(
        /\b(help|ajuda|socorro)\b.*\b(histórico|historico|nota|anot|interação|interacao|registrar?|obs)\b/
      )
    ) {
      await send(MSG.helpHistorico());
      return;
    }
    if (
      textoMinusculo.match(
        /\b(help|ajuda|socorro)\b.*\b(deletar?|excluir?|remover?|apagar?|delete)\b/
      )
    ) {
      await send(MSG.helpDeletar());
      return;
    }

    // ── Ajuda geral / menu ──
    if (
      textoMinusculo.match(
        /\b(ajuda|help|socorro|comandos|menu|o que você faz|o que voce faz|como usar|instruções|instrucoes|o que sabe|o que você sabe|quais comandos|que comandos)\b/
      )
    ) {
      await send(MSG.ajuda());
      return;
    }

    if (intencoesDetectadas.length === 0) {
      registrarLog(
        'INFO',
        CORRETOR_NOME,
        'INTENCAO_NAO_IDENTIFICADA',
        { texto: textoOriginal },
        'Nenhuma intenção mapeada encontrada. Exibindo menu de ajuda.'
      );
      await send(MSG.ajuda());
      return;
    }

    // ── DIRETIVA 4: Colisão de intenções ──
    // Exceção permitida: INTERACAO pode coexistir com ATUALIZAR (registrar + atualizar = contexto de histórico)
    const intencoesConflitantes = intencoesDetectadas.filter(
      (i) => !['INTERACAO', 'ATUALIZAR'].includes(i) || intencoesDetectadas.length > 2
    );
    const temColisaoReal =
      intencoesDetectadas.length >= 2 &&
      !(
        intencoesDetectadas.length === 2 &&
        intencoesDetectadas.includes('INTERACAO') &&
        intencoesDetectadas.includes('ATUALIZAR')
      );

    if (temColisaoReal) {
      registrarLog(
        'BLOQUEIO',
        CORRETOR_NOME,
        'COLISAO_INTENCOES',
        { intencoesDetectadas, texto: textoOriginal },
        `Bloqueio por colisão: ${intencoesDetectadas.join(' + ')} identificadas simultaneamente.`
      );
      await send(MSG.colisao(intencoesDetectadas));
      return;
    }

    // ── ROTEAMENTO DE INTENÇÃO ÚNICA ──
    const intencao = intencoesDetectadas[0];

    if (intencao === 'CRIAR') {
      await handleCriar(textoOriginal, textoMinusculo, send);
    } else if (intencao === 'LISTAR') {
      await handleListar(send);
    } else if (intencao === 'ATUALIZAR') {
      await handleAtualizar(textoOriginal, userId, send);
    } else if (intencao === 'DELETAR') {
      await handleDeletar(textoOriginal, userId, send);
    } else if (intencao === 'INTERACAO') {
      await handleInteracao(textoOriginal, userId, send);
    } else {
      await send(MSG.ajuda());
    }
  } catch (err) {
    // ── DIRETIVA 1: Captura de emergência do bloco principal ──
    try {
      registrarLog(
        'FATAL',
        'SISTEMA',
        'HANDLER_MESSAGE_EXCEPTION',
        { stack: err?.stack || String(err) },
        'Exceção capturada no nível máximo do handler de mensagens.'
      );
      await send(MSG.manutencao());
    } catch {
      // Se até o send falhar, registra no terminal e segue em frente. O bot NÃO cai.
      console.error(
        '💀 [FATAL CRÍTICO] Falha total no handler de mensagem. Bot mantido ativo.',
        err
      );
    }
  }
});

module.exports = app;
