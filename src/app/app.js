const { App } = require('@microsoft/teams.apps');
const { LocalStorage } = require('@microsoft/teams.common');

const storage = new LocalStorage();
const app = new App({ storage });

let bancoDadosCRM = [
  {
    id: 101,
    nome: 'João Silva',
    telefone: '(45) 99999-1111',
    empreendimento: 'Carmel',
    responsavel: 'Alex Wilber',
  },
  {
    id: 102,
    nome: 'João Silva',
    telefone: '(45) 98888-2222',
    empreendimento: 'Villa Bella',
    responsavel: 'Alex Wilber',
  },
  {
    id: 103,
    nome: 'Carlos Souza',
    telefone: '(45) 97777-3333',
    empreendimento: 'Carmel',
    responsavel: 'Mariana Costa',
  },
];

let memoriaConversa = {};

// === FUNÇÃO DE LOG EXIGIDA NO ESCOPO ===
function registrarLog(tipo, usuario, acao, dados) {
  const timestamp = new Date().toISOString();
  console.log(`\n================ [LOG SEGURO - ${tipo}] ===============`);
  console.log(`Data/Hora: ${timestamp}`);
  console.log(`Usuário do Teams: ${usuario}`);
  console.log(`Ação Detectada: ${acao}`);
  console.log(`Payload/Detalhes:`, JSON.stringify(dados, null, 2));
  console.log(`======================================================\n`);
}

app.on('message', async ({ send, activity }) => {
  const usuarioLogado = activity.from.name || 'Usuário Desconhecido';
  const userId = activity.from.id;
  const textoOriginal = activity.text.trim();
  const textoMinusculo = textoOriginal.toLowerCase();

  // === TRATANDO DUPLICIDADE DO JOÃO SILVA ===
  if (memoriaConversa[userId] && memoriaConversa[userId].acao === 'aguardando_escolha_joao') {
    if (textoOriginal === '1' || textoOriginal === '2') {
      const idEscolhido = textoOriginal === '1' ? 101 : 102;
      const leadSelecionado = bancoDadosCRM.find((l) => l.id === idEscolhido);

      // ETAPA 6: Montando Payload de Atualização simulado do CV CRM
      const payloadAtualizacao = {
        idlead: idEscolhido,
        contribuinte: usuarioLogado,
        interacao_origem: 'Microsoft Teams Bot',
        observacao: `Lead atualizado via comando de voz/texto no Teams por ${usuarioLogado}`,
      };

      registrarLog(
        'SUCESSO_INTEGRACAO',
        usuarioLogado,
        'ATUALIZAR_LEAD_RESOLVIDO',
        payloadAtualizacao
      );

      await send(
        `✅ [Compila na Fé] Perfeito! Você escolheu o **ID ${idEscolhido}**. Integração enviada para a API do CV CRM com sucesso!`
      );
      delete memoriaConversa[userId];
      return;
    } else {
      await send('⚠️ Por favor, responda apenas **1** ou **2**.');
      return;
    }
  }

  // === REGRA 1: CRIAR LEAD ===
  if (textoMinusculo.includes('crie') || textoMinusculo.includes('cadastre')) {
    const padraoTelefone = /(?:\(?\d{2}\)?\s?)?\d{4,5}-?\d{4}/g;
    const telefoneEncontrado = textoOriginal.match(padraoTelefone);

    if (!telefoneEncontrado) {
      registrarLog('BLOQUEIO_VALIDACAO', usuarioLogado, 'CRIAR_LEAD', {
        erro: 'Contato minimo ausente',
        texto: textoOriginal,
      });
      await send(
        `❌ [Compila na Fé] Erro: Não consegui identificar um número de telefone válido. Forneça o telefone para o cadastro.`
      );
      return;
    }

    let nomeLimpo = textoOriginal
      .replace(/(crie|cadastre|lead|telefone|fone)/gi, '')
      .split(',')[0]
      .trim();

    // ETAPA 6: Montando o Payload Oficial baseado na documentação do CV CRM (postlead)
    const payloadCvcrm = {
      nome: nomeLimpo || 'Novo Lead Anonimo',
      telefone: telefoneEncontrado[0],
      email: '',
      origem: 'Microsoft Teams Bot',
      gestor_responsavel: usuarioLogado, // Regra de governança: vira o dono!
      repassar: 'N',
    };

    // Salva na nossa lista local para simular o banco de dados atualizado
    bancoDadosCRM.push({ id: bancoDadosCRM.length + 101, ...payloadCvcrm });

    // Registra o log técnico no painel/terminal
    registrarLog('SUCESSO_INTEGRACAO', usuarioLogado, 'CRIAR_LEAD_API', payloadCvcrm);

    await send(
      `🎉 Lead **"${payloadCvcrm.nome}"** enviado para o CV CRM e vinculado a você (${usuarioLogado})!`
    );
    return;
  }

  // === REGRA 2: ATUALIZAR LEAD ===
  if (textoMinusculo.includes('atualize') || textoMinusculo.includes('edite')) {
    if (textoMinusculo.includes('joão') || textoMinusculo.includes('joao')) {
      memoriaConversa[userId] = { acao: 'aguardando_escolha_joao' };
      await send(
        `⚠️ Encontrei mais de um lead chamado "João Silva". Qual deles você deseja editar?\nDigite **1** para o ID 101 (Carmel)\nDigite **2** para o ID 102 (Villa Bella)`
      );
      return;
    }

    if (textoMinusculo.includes('carlos')) {
      const leadCarlos = bancoDadosCRM.find((l) => l.id === 103);
      if (leadCarlos.responsavel !== usuarioLogado) {
        // REGRA DE GOVERNANÇA: Tentativas bloqueadas devem gerar log!
        registrarLog('VIOLACAO_GOVERNANCA', usuarioLogado, 'ATUALIZAR_LEAD_NEGADO', {
          leadId: 103,
          donoAtual: leadCarlos.responsavel,
        });
        await send(
          `❌ Não foi possível atualizar o lead Carlos Souza. Esse lead pertence a ${leadCarlos.responsavel}. Acesso negado!`
        );
        return;
      }
    }
  }

  await send(
    `[Compila na Fé] Comandos de teste:\n- "Crie lead Roberto, fone 45988882222"\n- "Atualize o lead João Silva"\n- "Atualize o lead Carlos Souza"`
  );
});

module.exports = app;
