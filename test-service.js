require('dotenv').config();
const cvCrmService = require('./src/services/cvCrmService.js');

async function rodarTestes() {
  console.log('🧪 Iniciando bateria de testes do CV CRM Service...\n');
  console.log(
    `Modo Atual: ${process.env.USE_MOCK === 'true' ? '🟡 MOCK (Simulação)' : '🟢 API REAL'}\n`
  );

  try {
    // Teste 1: Criação de Lead
    console.log('--- TESTE 1: CRIAR LEAD ---');
    const mockLead = {
      nome: 'Teste Hackathon Biopark',
      telefone: '45999990000',
      origem: 'Script de Teste',
      gestor_responsavel: 'DevTester',
    };
    const leadCriado = await cvCrmService.criarLead(mockLead);
    console.log('✅ Sucesso! Resposta:', leadCriado, '\n');

    // Teste 2: Atualização de Lead
    console.log('--- TESTE 2: ATUALIZAR LEAD ---');
    const idParaTestar = leadCriado?.id || 101;
    const updateResult = await cvCrmService.atualizarLead(idParaTestar, {
      observacao: 'Teste de update',
    });
    console.log('✅ Sucesso! Resposta:', updateResult, '\n');

    // Teste 3: Registrar Interação
    console.log('--- TESTE 3: REGISTRAR INTERAÇÃO ---');
    const interacaoResult = await cvCrmService.registrarInteracao(idParaTestar, {
      descricao: 'Ligação de alinhamento feita.',
    });
    console.log('✅ Sucesso! Resposta:', interacaoResult, '\n');
  } catch (error) {
    console.error('❌ ERRO NO TESTE:', error.message);
  }
}

rodarTestes();
