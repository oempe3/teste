# Projeto Quadro de Disponibilidade - Pernambuco III

Sistema completo de monitoramento de equipamentos em tempo real com integração ao Google Sheets, geração automática de relatórios em PDF e sistema de autenticação seguro.

## 🚀 Funcionalidades Implementadas

### ✅ Sistema de Autenticação Seguro
- **Login com e-mail e senha** validados via Google Apps Script
- **Cadastro de usuários** com geração automática de senha
- **Proteção de páginas** via sessionStorage
- **Controle de permissões** (admin/operador)
- **Envio de credenciais por e-mail** automático
- **Sessão com expiração** de 24 horas

### ✅ Atualização de Equipamentos
- **Dropdown dinâmico** preenchido com dados do Google Sheets
- **Atualização específica de linhas** sem afetar outros dados
- **Formulário inteligente** que adapta campos baseado no status
- **Validação de dados** em tempo real
- **Histórico de modificações** visível na interface

### ✅ Upload e Substituição via CSV
- **Upload completo de planilha** via arquivo CSV
- **Validação de formato** e estrutura do arquivo
- **Substituição segura** de todos os dados
- **Feedback detalhado** do processo

### ✅ Geração Automática de PDF Diário
- **Identificação automática** de alterações do dia
- **HTML estilizado** para relatórios profissionais
- **Conversão automática** para PDF
- **Envio por e-mail** para lista de destinatários
- **Agendamento automático** via triggers do Google Apps Script

### ✅ Integração com Excel VBA
- **API exclusiva** para comunicação com Excel
- **Token de segurança** para autenticação
- **Código VBA completo** fornecido
- **Processamento em lote** de atualizações
- **Feedback de resultados** por linha

### ✅ Interface Moderna e Responsiva
- **Design inspirado em Flutter** com cores e componentes modernos
- **Cards compactos** para visualização máxima em uma página
- **Responsividade completa** para mobile e desktop
- **Horário de Brasília** em tempo real
- **Filtros avançados** por busca e status
- **Animações suaves** e micro-interações

## 📁 Estrutura do Projeto

```
status_p3-main/
├── index.html              # Página de login
├── cadastro.html           # Página de cadastro de usuários
├── entrada.html            # Formulário de atualização de equipamentos
├── status.html             # Painel de visualização de status
├── csv_update.html         # Upload e substituição via CSV
├── style.css               # Estilos CSS principais
├── auth.js                 # Sistema de autenticação
├── common.js               # Funções comuns
├── google-apps-script.js   # Código do Google Apps Script
├── excel-vba-integration.vba # Código VBA para Excel
├── logomarca.png           # Logomarca do sistema
├── favicon.png             # Ícone do site
└── README.md               # Este arquivo
```

## 🛠️ Configuração e Instalação

### 1. Configuração do Google Apps Script

1. Acesse [Google Apps Script](https://script.google.com/)
2. Crie um novo projeto
3. Cole o código do arquivo `google-apps-script.js`
4. Configure as variáveis:
   ```javascript
   const PLANILHA_EQUIPAMENTOS_ID = 'SEU_ID_DA_PLANILHA_EQUIPAMENTOS';
   const PLANILHA_USUARIOS_ID = 'SEU_ID_DA_PLANILHA_USUARIOS';
   ```
5. Publique como aplicação web:
   - Executar como: Eu
   - Quem tem acesso: Qualquer pessoa
6. Copie a URL do aplicativo web

### 2. Configuração das Planilhas Google Sheets

#### Planilha de Equipamentos
Crie uma planilha com as seguintes colunas:
```
TAG | STATUS | MOTIVO | PTS | OS | RETORNO | CADEADO | OBSERVACOES | MODIFICADO_POR | DATA
```

#### Planilha de Usuários
Crie uma planilha com as seguintes colunas:
```
NOME | EMAIL | SENHA | TIPO | AUTORIZADO | DATA_CADASTRO | DATA_ULTIMO_LOGIN
```

### 3. Configuração do Frontend

1. Atualize o endpoint em todos os arquivos JavaScript:
   ```javascript
   const endpoint = "SUA_URL_DO_GOOGLE_APPS_SCRIPT";
   ```

2. Para usar com Excel VBA, configure o token de segurança:
   ```javascript
   const SECURITY_TOKEN = "SEU_TOKEN_SEGURO_AQUI";
   ```

### 4. Configuração dos Gatilhos Automáticos

No Google Apps Script, execute a função `configurarGatilhos()` para:
- Configurar envio diário de PDF às 18:00
- Configurar limpeza automática de sessões expiradas

## 📧 Configuração de E-mails

### E-mails de Cadastro
Os e-mails são enviados automaticamente via `MailApp.sendEmail()` quando um usuário se cadastra.

### E-mails de Relatório Diário
Configure a lista de destinatários no Google Apps Script:
```javascript
const destinatarios = [
  'admin@pernambuco3.com.br',
  'operacao@pernambuco3.com.br'
  // Adicionar mais e-mails conforme necessário
];
```

## 🔧 Integração com Excel VBA

### Configuração no Excel

1. Abra o Excel e pressione `Alt + F11` para abrir o VBA
2. Insira um novo módulo
3. Cole o código do arquivo `excel-vba-integration.vba`
4. Configure as constantes:
   ```vba
   Const API_ENDPOINT As String = "SUA_URL_DO_GOOGLE_APPS_SCRIPT"
   Const SECURITY_TOKEN As String = "SEU_TOKEN_SEGURO_AQUI"
   ```

### Uso do VBA

1. Execute `CriarTemplatePlanilha()` para criar um template
2. Preencha os dados na planilha
3. Execute `AtualizarEquipamentoViaExcel()` para enviar os dados

## 🎨 Personalização

### Cores e Temas
As cores principais estão definidas no `:root` do CSS:
```css
:root {
  --primary-color: #2563eb;
  --success-color: #10b981;
  --warning-color: #f59e0b;
  --error-color: #ef4444;
  /* ... */
}
```

### Logomarca
Substitua o arquivo `logomarca.png` pela logomarca desejada (recomendado: 200x80px).

### Favicon
Substitua o arquivo `favicon.png` pelo ícone desejado (recomendado: 32x32px).

## 📱 Responsividade

O sistema é totalmente responsivo e funciona em:
- **Desktop** (1200px+)
- **Tablet** (768px - 1199px)
- **Mobile** (até 767px)

### Características Mobile
- Cards de equipamentos empilhados
- Menu de navegação adaptado
- Formulários otimizados para toque
- Filtros colapsáveis

## 🔒 Segurança

### Autenticação
- Senhas armazenadas em texto plano (recomenda-se hash em produção)
- Sessões com expiração automática
- Validação de permissões por página

### API
- Token de segurança para integração VBA
- Validação de tipos de requisição
- Logs de erro detalhados

## 📊 Relatórios

### PDF Diário
- Gerado automaticamente às 18:00
- Inclui apenas alterações do dia
- Enviado por e-mail para lista configurada
- Design profissional com estatísticas

### Formato do Relatório
- Cabeçalho com data e logo
- Resumo estatístico
- Detalhes de cada alteração
- Rodapé com timestamp

## 🚀 Publicação

### GitHub Pages
1. Faça upload dos arquivos para um repositório GitHub
2. Ative o GitHub Pages nas configurações
3. O site estará disponível em `https://usuario.github.io/repositorio`

### Servidor Web
1. Faça upload dos arquivos para seu servidor
2. Configure HTTPS (recomendado)
3. Teste todas as funcionalidades

## 🐛 Solução de Problemas

### Erro de CORS
Se houver problemas de CORS, verifique:
- URL do Google Apps Script está correta
- Aplicação web está publicada corretamente
- Permissões estão configuradas como "Qualquer pessoa"

### E-mails não enviados
Verifique:
- Permissões do Gmail no Google Apps Script
- Lista de destinatários está correta
- Cota de e-mails do Google não foi excedida

### Dados não carregam
Verifique:
- IDs das planilhas estão corretos
- Planilhas têm as colunas corretas
- Permissões de leitura estão configuradas

## 📞 Suporte

Para suporte técnico ou dúvidas sobre implementação:
- Verifique os logs do Google Apps Script
- Teste as funções individualmente
- Consulte a documentação do Google Apps Script

## 📄 Licença

Este projeto foi desenvolvido para uso interno da Pernambuco III. Todos os direitos reservados.

---

**Desenvolvido com ❤️ para Pernambuco III**

