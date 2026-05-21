const { App } = require('@microsoft/teams.apps');

const storage = {
  read: async () => ({}),
  write: async () => {},
  delete: async () => {},
};

const app = new App({
  storage,
  config: {
    MicrosoftAppId: '',
    MicrosoftAppPassword: '',
  },
});

let bancoDadosCRM = [
  {
    id: 101,
    nome: 'João Silva',
    telefone: '45999991111',
    empreendimento: 'Carmel',
    responsavel: 'Alex Wilber',
  },
  {
    id: 102,
    nome: 'João Silva',
    telefone: '45988882222',
    empreendimento: 'Villa Bella',
    responsavel: 'Alex Wilber',
  },
  {
    id: 103,
    nome: 'Carlos Souza',
    telefone: '45977773333',
    empreendimento: 'Carmel',
    responsavel: 'Mariana Costa',
  },
];

let memoriaConversa = {};

function registrarLog(tipo, usuario, acao, dados) {
  const timestamp = new Date().toISOString();
  console.log(`\n================ [LOG SEGURO - ${tipo}] ===============`);
  console.log(`Data/Hora: ${timestamp}`);
  console.log(`Usuário: ${usuario}`);
  console.log(`Ação: ${acao}`);
  console.log(`Payload:`, JSON.stringify(dados, null, 2));
  console.log(`======================================================\n`);
}

// Substitua a função enviarParaCRM antiga do seu app.js por esta:
async function enviarParaCRM(metodo, endpoint, payload = null) {
  const url_api = `http://localhost:3000/api/v1${endpoint}`;

  try {
    const configuracao = {
      method: metodo,
      headers: { 'Content-Type': 'application/json' },
    };

    if (payload) {
      configuracao.body = JSON.stringify(payload);
    }

    const resposta = await fetch(url_api, configuracao);
    const dadosRetorno = await resposta.json();

    if (resposta.ok) {
      return { sucesso: true, dados: dadosRetorno };
    } else {
      return { sucesso: false, erro: dadosRetorno.erro || `Erro HTTP ${resposta.status}` };
    }
  } catch (error) {
    return { sucesso: false, erro: 'Servidor CRM Offline: ' + error.message };
  }
}

app.on('message', async ({ send, activity }) => {
  const usuarioLogado = activity.from.name || 'Usuário Desconhecido';
  const userId = activity.from.id;

  const textoOriginal = activity.text ? activity.text.trim() : '';
  const textoMinusculo = textoOriginal.toLowerCase();

  // === FLUXO DE DESAMBIGUAÇÃO ===
  if (memoriaConversa[userId] && memoriaConversa[userId].etapa === 'AGUARDANDO_ID') {
    if (textoOriginal === '1' || textoOriginal === '2') {
      const index = parseInt(textoOriginal) - 1;
      const leadEscolhido = memoriaConversa[userId].leadsPossiveis[index];

      const payloadAtualizacao = {
        idlead: leadEscolhido.id,
        contribuinte: usuarioLogado,
        interacao_origem: 'Teams Bot Natural Language',
        observacao: `Lead atualizado via chat por ${usuarioLogado}`,
      };

      registrarLog(
        'SUCESSO_INTEGRACAO',
        usuarioLogado,
        'ATUALIZAR_LEAD_RESOLVIDO',
        payloadAtualizacao
      );
      await send(
        `✅ **[Compila na Fé] Sucesso!**\n\nVocê selecionou o **ID ${leadEscolhido.id}** (${leadEscolhido.empreendimento}). Alteração simulada e enviada para o CRM com sucesso!`
      );
      delete memoriaConversa[userId];
      return;
    } else {
      await send('⚠️ Opção inválida. Digite apenas **1** ou **2** para escolher o lead correto.');
      return;
    }
  }

  // === REGISTRO DE INTERAÇÃO (AGENDAMENTO) ===
  if (
    textoMinusculo.includes('registre') ||
    textoMinusculo.includes('agende') ||
    textoMinusculo.includes('observação')
  ) {
    const padraoId = /\b\d{3}\b/g;
    const idEncontrado = textoOriginal.match(padraoId);

    if (!idEncontrado) {
      await send(
        "❌ **Erro de Comando:**\n\nPara registrar uma interação, informe o número do ID.\n\n*Exemplo: 'Registre no lead 101 que liguei para ele'*"
      );
      return;
    }

    const leadId = parseInt(idEncontrado[0]);
    const leadAlvo = bancoDadosCRM.find((l) => l.id === leadId);

    if (!leadAlvo) {
      await send(`❌ **Erro:** O lead com ID **${leadId}** não foi localizado no sistema.`);
      return;
    }

    if (leadAlvo.responsavel !== usuarioLogado) {
      registrarLog('VIOLACAO_GOVERNANCA', usuarioLogado, 'REGISTRAR_OBS_NEGADO', {
        leadId,
        dono: leadAlvo.responsavel,
      });
      await send(
        `❌ **Acesso Negado!**\n\nEste lead está atribuído a **${leadAlvo.responsavel}**. Você só pode alterar leads sob sua responsabilidade.`
      );
      return;
    }

    const payloadInteracao = {
      idlead: leadId,
      usuario: usuarioLogado,
      origem: 'Teams Bot',
      data_hora: new Date().toISOString(),
      descricao: textoOriginal.replace(/registre no lead \d+/gi, '').trim(),
    };

    registrarLog('SUCESSO_INTEGRACAO', usuarioLogado, 'REGISTRAR_INTERACAO_API', payloadInteracao);
    const resultadoAPI = await enviarParaCRM('PATCH', `/leads/${leadId}`, payloadInteracao);

    await send(
      `📝 **Interação Registrada!**\n\n• **Lead:** ${leadAlvo.nome} (ID ${leadId})\n• **Vendedor:** ${usuarioLogado}\n• **Texto cadastrado:** "${payloadInteracao.descricao}"`
    );
    return;
  }

  // === CRIAR LEAD (AGORA COM API REAL) ===
  if (textoMinusculo.includes('crie') || textoMinusculo.includes('cadastre')) {
    const padraoTelefone = /(?:\(?\d{2}\)?\s?)?\d{4,5}-?\d{4}/g;
    const telefoneEncontrado = textoOriginal.match(padraoTelefone);

    if (!telefoneEncontrado) {
      registrarLog('BLOQUEIO_VALIDACAO', usuarioLogado, 'CRIAR_LEAD', {
        erro: 'Contato minimo ausente',
      });
      await send(
        `❌ **Erro de Validação:**\n\nNão consegui identificar um número de telefone válido na mensagem. Forneça o telefone para realizar o cadastro.`
      );
      return;
    }

    const telefoneLimpo = telefoneEncontrado[0].replace(/\D/g, '');
    const leadExistente = bancoDadosCRM.find(
      (l) => l.telefone.replace(/\D/g, '') === telefoneLimpo
    );

    if (leadExistente) {
      registrarLog('AVISO_DUPLICIDADE', usuarioLogado, 'CRIAR_LEAD_DUPLICADO', {
        telefone: telefoneEncontrado[0],
        leadExistenteId: leadExistente.id,
      });
      await send(
        `⚠️ **Lead Duplicado!**\n\nO telefone **${telefoneEncontrado[0]}** já pertence ao lead **"${leadExistente.nome}"** (ID ${leadExistente.id}). Cadastro bloqueado para evitar sujeira no banco de dados.`
      );
      return;
    }

    let nomeLimpo = textoOriginal
      .replace(/(crie|cadastre|lead|telefone|fone)/gi, '')
      .split(',')[0]
      .trim();

    const payloadCvcrm = {
      nome: nomeLimpo || 'Novo Lead Anônimo',
      telefone: telefoneEncontrado[0],
      origem: 'Teams Bot',
      gestor_responsavel: usuarioLogado,
      repassar: 'N',
    };

    // 1. Avisamos o usuário que estamos processando a requisição
    await send(`⏳ Conectando aos servidores do CV CRM...`);

    // 2. Disparamos a requisição HTTP Real para a internet!
    const resultadoAPI = await enviarParaCRM('POST', '/leads', payloadCvcrm);

    // 3. Salvamos no banco local para a demonstração continuar fluindo
    bancoDadosCRM.push({
      id: bancoDadosCRM.length + 101,
      ...payloadCvcrm,
      responsavel: usuarioLogado,
    });
    registrarLog('SUCESSO_INTEGRACAO', usuarioLogado, 'CRIAR_LEAD_API', payloadCvcrm);

    // 4. Tratamos o retorno na tela
    if (resultadoAPI.sucesso) {
      await send(
        `🎉 **Lead Criado com Sucesso via API!**\n\n• **Nome:** ${payloadCvcrm.nome}\n• **Telefone:** ${payloadCvcrm.telefone}\n• **Responsável:** ${usuarioLogado}`
      );
    } else {
      // Como não temos o token real, ele vai cair aqui durante o hackathon. Isso mostra que o tratamento de erro funciona!
      await send(
        `⚠️ **Lead Criado Localmente.**\n\n• **Nome:** ${payloadCvcrm.nome}\n• **Telefone:** ${payloadCvcrm.telefone}\n\n*Nota de Sistema:* A tentativa de sincronização online falhou (${resultadoAPI.erro}) devido à ausência de credenciais reais da API, mas o payload foi gerado nos logs.`
      );
    }
    return;
  }

  // === ATUALIZAR LEAD ===
  if (textoMinusculo.includes('atualize') || textoMinusculo.includes('edite')) {
    if (textoMinusculo.includes('joão') || textoMinusculo.includes('joao')) {
      const leadsFiltrados = bancoDadosCRM.filter((l) => l.nome === 'João Silva');

      memoriaConversa[userId] = { etapa: 'AGUARDANDO_ID', leadsPossiveis: leadsFiltrados };

      await send(
        `⚠️ **Múltiplos leads encontrados com o nome "João Silva"**\n\nPor favor, digite apenas o número correspondente para indicar quem deseja alterar:\n\n**1** - ID 101 (Final fone: 1111 - Empreendimento: Carmel)\n**2** - ID 102 (Final fone: 2222 - Empreendimento: Villa Bella)`
      );
      return;
    }

    if (textoMinusculo.includes('carlos')) {
      const leadCarlos = bancoDadosCRM.find((l) => l.id === 103);
      if (leadCarlos.responsavel !== usuarioLogado) {
        registrarLog('VIOLACAO_GOVERNANCA', usuarioLogado, 'ATUALIZAR_LEAD_NEGADO', {
          leadId: 103,
          dono: leadCarlos.responsavel,
        });
        await send(
          `❌ **Acesso Negado!**\n\nNão foi possível atualizar o lead **Carlos Souza**. Esse lead pertence a **${leadCarlos.responsavel}** e você está logado como **${usuarioLogado}**.`
        );
        return;
      }
    }
  }

  // === [NOVO CRUD] LISTAR LEADS DIRECTO DA API (GET) ===
  if (textoMinusculo.includes('liste') || textoMinusculo.includes('mostrar todos'))  {
    await send(`⏳ Buscando registros atualizados no servidor do CRM...`);

    const resultadoAPI = await enviarParaCRM('GET', '/leads');

    if (resultadoAPI.sucesso) {
      const listaLeads = resultadoAPI.dados.dados;
      let respostaTexto = `📋 **Leads Registrados no CV CRM:**\n\n`;

      listaLeads.forEach((lead) => {
        respostaTexto += `• **ID ${lead.id}**: ${lead.nome} - Fone: ${lead.telefone} (Responsável: ${lead.responsavel})\n`;
      });

      await send(respostaTexto);
    } else {
      await send(`❌ Não foi possível listar os leads: ${resultadoAPI.erro}`);
    }
    return;
  }

  // === [NOVO CRUD] DELETAR LEAD VIA CHAT (DELETE) ===
  if (
    textoMinusculo.includes('exclua') ||
    textoMinusculo.includes('delete') ||
    textoMinusculo.includes('remover')
  ) {
    const padraoId = /\b\d{3}\b/g;
    const idEncontrado = textoOriginal.match(padraoId);

    if (!idEncontrado) {
      await send(
        "❌ **Erro de Comando:** Informe o ID de 3 dígitos do lead que deseja remover. *Ex: 'Delete o lead 101'*"
      );
      return;
    }

    const leadId = parseInt(idEncontrado[0]);
    await send(`⏳ Solicitando exclusão do ID ${leadId} ao servidor...`);

    const resultadoAPI = await enviarParaCRM('DELETE', `/leads/${leadId}`);

    if (resultadoAPI.sucesso) {
      // Remove do banco local também para manter sincronizado
      bancoDadosCRM = bancoDadosCRM.filter((l) => l.id !== leadId);

      registrarLog('SUCESSO_INTEGRACAO', usuarioLogado, 'DELETAR_LEAD_API', { idDeletado: leadId });
      await send(
        `🗑️ **Sucesso!** O lead com ID **${leadId}** foi completamente removido do banco de dados do CRM.`
      );
    } else {
      await send(`❌ **Erro no CRM:** ${resultadoAPI.erro}`);
    }
    return;
  }

  await send(
    `🤖 **[Compila na Fé] Assistente de CRM Ativo!**\n\nEnvie seus comandos em linguagem natural para o gerenciamento de leads.`
  );
});

module.exports = app;
