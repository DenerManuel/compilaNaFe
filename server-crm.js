const express = require('express');
const app = express();
const PORT = 3000; // A API vai rodar na porta 3000

app.use(express.json());

// Nosso Banco de Dados simulado dentro da API do CRM
let bancoCRM = [
  {
    id: 101,
    nome: 'João Silva',
    telefone: '45999991111',
    empreendimento: 'Carmel',
    responsavel: 'Alex Wilber',
    interacoes: [],
  },
  {
    id: 102,
    nome: 'João Silva',
    telefone: '45988882222',
    empreendimento: 'Villa Bella',
    responsavel: 'Alex Wilber',
    interacoes: [],
  },
  {
    id: 103,
    nome: 'Carlos Souza',
    telefone: '45977773333',
    empreendimento: 'Carmel',
    responsavel: 'Mariana Costa',
    interacoes: [],
  },
];

// [READ] - Rota para listar todos os leads no CRM
app.get('/api/v1/leads', (req, res) => {
  console.log('➡️ [CRM API] GET: Listando todos os leads');
  res.json({ sucesso: true, dados: bancoCRM });
});

// [CREATE] - Rota para criar um novo lead
app.post('/api/v1/leads', (req, res) => {
  const { nome, telefone, origem, gestor_responsavel } = req.body;

  console.log('➡️ [CRM API] POST: Recebido payload de criação:', req.body);

  const novoLead = {
    id: bancoCRM.length + 101,
    nome: nome || 'Novo Lead',
    telefone: telefone,
    empreendimento: 'Não Informado',
    responsavel: gestor_responsavel || 'Sem Atribuição',
    interacoes: [],
  };

  bancoCRM.push(novoLead);
  res
    .status(201)
    .json({ sucesso: true, mensagem: 'Lead cadastrado com sucesso no CRM!', lead: novoLead });
});

// [UPDATE] - Rota para atualizar ou registrar interações em um lead específico
app.patch('/api/v1/leads/:id', (req, res) => {
  const idAlvo = parseInt(req.params.id);
  console.log(`➡️ [CRM API] PATCH: Atualizando lead ID ${idAlvo}. Dados recebidos:`, req.body);

  const lead = bancoCRM.find((l) => l.id === idAlvo);

  if (!lead) {
    return res.status(404).json({ sucesso: false, erro: 'Lead não localizado no banco do CRM.' });
  }

  // Se o payload conter uma descrição, salvamos como interação histórica
  if (req.body.descricao) {
    lead.interacoes.push({
      data: req.body.data_hora || new Date().toISOString(),
      texto: req.body.descricao,
      autor: req.body.usuario || 'Sistema',
    });
    return res.json({ sucesso: true, mensagem: 'Interação gravada no histórico do CRM!', lead });
  }

  // Atualização de campos comuns se enviados
  if (req.body.nome) lead.nome = req.body.nome;
  if (req.body.telefone) lead.telefone = req.body.telefone;

  res.json({ sucesso: true, mensagem: 'Dados do lead atualizados no CRM!', lead });
});

// [DELETE] - Rota para deletar um lead
app.delete('/api/v1/leads/:id', (req, res) => {
  const idAlvo = parseInt(req.params.id);
  console.log(`➡️ [CRM API] DELETE: Removendo lead ID ${idAlvo}`);

  const index = bancoCRM.findIndex((l) => l.id === idAlvo);

  if (index === -1) {
    return res.status(404).json({ sucesso: false, erro: 'Lead não existe no CRM.' });
  }

  bancoCRM.splice(index, 1);
  res.json({ sucesso: true, mensagem: `Lead ${idAlvo} deletado com sucesso do CRM.` });
});

app.listen(PORT, () => {
  console.log(`\n======================================================`);
  console.log(`🚀 API FAKE DO CRM RODANDO EM: http://localhost:${PORT}`);
  console.log(`======================================================\n`);
});
