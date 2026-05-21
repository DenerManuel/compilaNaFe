# Visão geral do modelo de Chatbot de IA Básico

Este modelo de aplicativo é desenvolvido sobre o [Microsoft Teams SDK](https://aka.ms/teams-ai-library-v2).
Ele demonstra um aplicativo de agente que responde a perguntas de usuários como o ChatGPT, o que permite que seus usuários conversem com o agente de IA no Teams.

## Começando com o modelo

> **Pré-requisitos**
>
> Para executar o modelo em sua máquina de desenvolvimento local, você precisará de:
>
> - [Node.js](https://nodejs.org/), versões suportadas: 20, 22.
> - [Extensão do Microsoft 365 Agents Toolkit para Visual Studio Code](https://aka.ms/teams-toolkit) (versão mais recente) ou [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli).
> - Preparar o seu próprio recurso do [Azure OpenAI](https://aka.ms/oai/access).

> Para depuração local usando o Microsoft 365 Agents Toolkit CLI, você precisa seguir algumas etapas extras descritas em [Configurar o Microsoft 365 Agents Toolkit CLI para depuração local](https://aka.ms/teamsfx-cli-debugging).

1. Primeiro, selecione o ícone do Microsoft 365 Agents Toolkit à esquerda na barra de ferramentas do VS Code.
2. No arquivo *env/.env.playground.user*, preencha a sua chave do Azure OpenAI em `SECRET_AZURE_OPENAI_API_KEY=<sua-chave>`, o endpoint em `AZURE_OPENAI_ENDPOINT=<seu-endpoint>` e o nome da implantação em `AZURE_OPENAI_DEPLOYMENT_NAME=<sua-implantacao>`.
3. Pressione F5 para iniciar a depuração, o que inicia o seu aplicativo no Microsoft 365 Agents Playground usando um navegador da web. Selecione `Debug in Microsoft 365 Agents Playground`.
4. Você pode enviar qualquer mensagem para obter uma resposta do agente.

**Parabéns**! Você está executando um aplicativo que agora pode interagir com os usuários no Microsoft 365 Agents Playground:

![agente de chat de ia](https://github.com/user-attachments/assets/984af126-222b-4c98-9578-0744790b103a)

## O que está incluído no modelo

| Pasta | Conteúdo |
| - | - |
| `.vscode` | Arquivos do VSCode para depuração |
| `appPackage` | Modelos para o manifesto do aplicativo |
| `env` | Arquivos de ambiente |
| `infra` | Modelos para provisionamento de recursos do Azure |
| `src` | O código-fonte do aplicativo |

Os arquivos a seguir podem ser personalizados e demonstram um exemplo de implementação para você começar.

| Arquivo | Conteúdo |
| - | - |
|`src/index.js`| Ponto de entrada do aplicativo. |
|`src/config.js`| Define as variáveis de ambiente. |
|`src/app/instructions.txt`| Define o prompt. |
|`src/app/app.js`| Lida com as lógicas de negócios para o Chatbot de IA Básico. |

A seguir estão os arquivos de projeto específicos do Microsoft 365 Agents Toolkit. Você pode [visitar um guia completo no Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) para entender como o Microsoft 365 Agents Toolkit funciona.

| Arquivo | Conteúdo |
| - | - |
|`m365agents.yml`| Este é o arquivo de projeto principal do Microsoft 365 Agents Toolkit. O arquivo de projeto define duas coisas principais: Propriedades e configurações e Definições de estágio. |
|`m365agents.local.yml`| Ele substitui o `m365agents.yml` com ações que habilitam a execução e depuração locais. |
|`m365agents.playground.yml`| Ele substitui o `m365agents.yml` com ações que habilitam a execução e depuração locais no Microsoft 365 Agents Playground. |

## Estenda o modelo

Para estender o modelo de Chatbot de IA Básico com mais recursos de IA, explore a [documentação do Microsoft Teams SDK](https://aka.ms/m365-agents-toolkit/teams-agent-extend-ai).

## Informações adicionais e referências

- [Documentação do Microsoft 365 Agents Toolkit](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Exemplos do Microsoft 365 Agents Toolkit](https://github.com/OfficeDev/TeamsFx-Samples)
