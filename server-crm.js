const express = require('express');
const app = express();
const PORT = 4000;

app.use(express.json());

// Middleware para simular um pequeno delay de rede (deixa a apresentação mais realista)
app.use((req, res, next) => {
  setTimeout(next, 600);
});

// Banco de Dados Mockado (Estrutura idêntica ao CV CRM)
let bancoCRM = [
  {
    id: 80,
    nome: 'João Silva',
    telefone: '45999991111',
    email: 'joao.silva101@teste.com',
    idimobiliaria: 2,
    imobiliaria_nome: 'Pratti',
    idcorretor: 464,
    corretor_nome: 'Leonardo',
    idsituacao: 3,
    interacoes: [
      { data: new Date().toISOString(), texto: 'Primeiro contato via WhatsApp', autor: 'Leonardo' },
    ],
  },
  {
    id: 81,
    nome: 'João Silva', // Nome duplicado propositalmente para demonstrar a Regra 5 (Ambiguidade)
    telefone: '45988882222',
    email: 'joao.silva102@teste.com',
    idimobiliaria: 2,
    imobiliaria_nome: 'Pratti',
    idcorretor: 120,
    corretor_nome: 'Amanda',
    idsituacao: 1,
    interacoes: [],
  },
  {
    id: 82,
    nome: 'Carlos Souza',
    telefone: '45977773333',
    email: 'carlos@souza.com',
    idimobiliaria: 2,
    imobiliaria_nome: 'Pratti',
    idcorretor: 464,
    corretor_nome: 'Leonardo',
    idsituacao: 3,
    interacoes: [],
  },
];

// ==========================================
// [R] - READ: Listar Leads
// ==========================================
app.get('/api/v1/comercial/leads', (req, res) => {
  console.log('➡️ [MOCK API] GET: Listando leads na base.');

  // O CV CRM retorna a lista dentro de um objeto { dados: [...] }
  res.status(200).json({
    sucesso: true,
    dados: bancoCRM,
    quantidade: bancoCRM.length,
  });
});

// ==========================================
// [C] - CREATE: Criar Novo Lead
// ==========================================
app.post('/api/v1/comercial/leads', (req, res) => {
  const { nome, telefone, email, gestor, imobiliaria, corretor } = req.body;
  console.log(`➡️ [MOCK API] POST: Tentativa de cadastro do lead "${nome}"`);

  // Simulação de trava do CRM: Rejeita se faltar dados obrigatórios
  if (!nome || !telefone || !email) {
    return res.status(400).json({
      sucesso: false,
      codigo: 400,
      mensagem: 'Parâmetros obrigatórios ausentes. Verifique nome, telefone e email.',
    });
  }

  // Simulação de erro 400 por duplicidade (Caso a checagem do bot falhe)
  const duplicado = bancoCRM.find((l) => l.email === email || l.telefone === telefone);
  if (duplicado) {
    return res.status(400).json({
      sucesso: false,
      codigo: 400,
      mensagem: `Lead já cadastrado com este e-mail ou telefone. ID associado: ${duplicado.id}`,
    });
  }

  const novoId = bancoCRM.length > 0 ? Math.max(...bancoCRM.map((l) => l.id)) + 1 : 1;

  const novoLead = {
    id: novoId,
    nome,
    telefone,
    email,
    idimobiliaria: imobiliaria?.id || 2,
    imobiliaria_nome: 'Pratti',
    idcorretor: corretor?.id || 464,
    corretor_nome: 'Leonardo', // Mockado para o fallback logado
    idsituacao: 3,
    interacoes: [],
  };

  bancoCRM.push(novoLead);

  // Retorno exato da documentação do CV CRM
  res.status(200).json({
    sucesso: true,
    id: novoId.toString(),
    idimobiliaria: novoLead.idimobiliaria.toString(),
    imobiliaria_nome: novoLead.imobiliaria_nome,
    idcorretor: novoLead.idcorretor.toString(),
    corretor_nome: novoLead.corretor_nome,
    idsituacao: '3',
    mensagem: 'Lead cadastrado com sucesso',
    codigo: 200,
  });
});

// ==========================================
// [U] - UPDATE: Atualizar ou Adicionar Histórico
// ==========================================
app.patch('/api/v1/comercial/leads/:id', (req, res) => {
  const idAlvo = parseInt(req.params.id);
  console.log(`➡️ [MOCK API] PATCH: Recebido payload para ID ${idAlvo}`, req.body);

  const leadIndex = bancoCRM.findIndex((l) => l.id === idAlvo);

  if (leadIndex === -1) {
    return res.status(404).json({
      sucesso: false,
      codigo: 404,
      mensagem: 'Lead não encontrado.',
    });
  }

  // Se o payload contém "descricao", é uma anotação de histórico
  if (req.body.descricao) {
    bancoCRM[leadIndex].interacoes.push({
      data: req.body.data_hora || new Date().toISOString(),
      texto: req.body.descricao,
      autor: req.body.usuario || 'Sistema Bot',
    });

    return res.status(200).json({
      sucesso: true,
      mensagem: 'Histórico inserido com sucesso',
      codigo: 200,
    });
  }

  // Atualização padrão de campos de cadastro
  if (req.body.nome) bancoCRM[leadIndex].nome = req.body.nome;
  if (req.body.telefone) bancoCRM[leadIndex].telefone = req.body.telefone;
  if (req.body.email) bancoCRM[leadIndex].email = req.body.email;

  res.status(200).json({
    sucesso: true,
    mensagem: 'Lead atualizado com sucesso',
    codigo: 200,
  });
});

// ==========================================
// [D] - DELETE: Remover Lead
// ==========================================
app.delete('/api/v1/comercial/leads/:id', (req, res) => {
  const idAlvo = parseInt(req.params.id);
  console.log(`➡️ [MOCK API] DELETE: Solicitação de exclusão ID ${idAlvo}`);

  const leadIndex = bancoCRM.findIndex((l) => l.id === idAlvo);

  if (leadIndex === -1) {
    return res.status(404).json({
      sucesso: false,
      codigo: 404,
      mensagem: 'Lead não encontrado para exclusão.',
    });
  }

  bancoCRM.splice(leadIndex, 1);

  res.status(200).json({
    sucesso: true,
    mensagem: 'Lead removido permanentemente.',
    codigo: 200,
  });
});

// Inicialização do Servidor
app.listen(PORT, () => {
  console.log(`\n======================================================`);
  console.log(`🚀 [PRATTI] MOCK API DO CV CRM INICIADA`);
  console.log(`🔗 Endpoint Base: http://localhost:${PORT}/api/v1/comercial/leads`);
  console.log(`======================================================\n`);
});
