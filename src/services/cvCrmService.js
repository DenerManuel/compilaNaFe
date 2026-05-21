const axios = require('axios');

class CVCrmService {
  constructor() {
    // 1. Lemos a Feature Flag do ambiente
    this.useMock = process.env.USE_MOCK === 'true';

    this.client = axios.create({
      baseURL: process.env.CV_CRM_BASE_URL,
      timeout: parseInt(process.env.API_TIMEOUT) || 5000,
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${process.env.CV_CRM_TOKEN}`,
      },
    });
  }

  _handleError(error, contextAction) {
    console.error(`[ERRO INTEGRAÇÃO - ${contextAction}]:`, error.message);
    if (error.response) {
      if (error.response.status === 401 || error.response.status === 403)
        throw new Error('Falha de autenticação.');
      if (error.response.status >= 500) throw new Error('O servidor do CRM está instável.');
      throw new Error(`Erro no CRM: ${error.response.data.mensagem || 'Validação recusada.'}`);
    } else if (error.request) {
      throw new Error('Tempo limite esgotado. O CRM parece estar indisponível.');
    } else {
      throw new Error('Erro interno ao preparar a requisição.');
    }
  }

  /**
   * Utilitário interno para simular o tempo de latência de uma API real
   */
  _simularLatencia(ms = 800) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  async criarLead(leadData) {
    // === O BOTÃO DE PÂNICO EM AÇÃO ===
    if (this.useMock) {
      console.log('🟡 [MOCK MODE] Simulando criação de lead no CV CRM...');
      await this._simularLatencia(); // Pausa o código para parecer uma requisição real
      return { id: Math.floor(Math.random() * 9000) + 1000 }; // Retorna um ID fake aleatório
    }
    // ==================================

    try {
      const payload = {
        nome: leadData.nome,
        telefone: leadData.telefone,
        origem: leadData.origem,
        idusuario: leadData.gestor_responsavel,
        repassar: 'N',
      };
      const response = await this.client.post('/lead', payload);
      return response.data;
    } catch (error) {
      this._handleError(error, 'CRIAR_LEAD');
    }
  }

  async atualizarLead(leadId, updateData) {
    if (this.useMock) {
      console.log(`🟡 [MOCK MODE] Simulando atualização do lead ${leadId}...`);
      await this._simularLatencia();
      return { sucesso: true, id: leadId };
    }

    try {
      const response = await this.client.put(`/lead/${leadId}`, updateData);
      return response.data;
    } catch (error) {
      this._handleError(error, `ATUALIZAR_LEAD_${leadId}`);
    }
  }

  async registrarInteracao(leadId, interacaoData) {
    if (this.useMock) {
      console.log(`🟡 [MOCK MODE] Simulando registro de interação no lead ${leadId}...`);
      await this._simularLatencia();
      return { sucesso: true, id_interacao: Math.floor(Math.random() * 9000) };
    }

    try {
      const payload = {
        idlead: leadId,
        descricao: interacaoData.descricao,
        tipo: 'Teams Bot Note',
        data: new Date().toISOString(),
      };
      const response = await this.client.post(`/lead/${leadId}/interacao`, payload);
      return response.data;
    } catch (error) {
      this._handleError(error, `REGISTRAR_INTERACAO_${leadId}`);
    }
  }
}

module.exports = new CVCrmService();
