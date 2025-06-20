// Google Apps Script - Sistema de Monitoramento de Equipamentos
// Este código deve ser copiado para o Google Apps Script

// IDs das planilhas (substitua pelos IDs reais)
const PLANILHA_EQUIPAMENTOS_ID = 'SEU_ID_DA_PLANILHA_EQUIPAMENTOS';
const PLANILHA_USUARIOS_ID = 'SEU_ID_DA_PLANILHA_USUARIOS';

// Função principal que processa todas as requisições
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    const tipo = dados.type;
    
    switch (tipo) {
      case 'login':
        return ContentService.createTextOutput(JSON.stringify(processarLogin(dados)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'cadastro':
        return ContentService.createTextOutput(JSON.stringify(processarCadastro(dados)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'update_row':
        return ContentService.createTextOutput(JSON.stringify(atualizarEquipamento(dados)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'atualizacaoExcel':
        return ContentService.createTextOutput(JSON.stringify(atualizarViaExcel(dados)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'full_replace':
        return ContentService.createTextOutput(JSON.stringify(substituirPlanilhaCompleta(dados)))
          .setMimeType(ContentService.MimeType.JSON);
      
      default:
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'Tipo de operação não reconhecido'
        })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    console.error('Erro no doPost:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'Erro interno do servidor: ' + error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Função GET para ler dados dos equipamentos
function doGet(e) {
  try {
    const planilha = SpreadsheetApp.openById(PLANILHA_EQUIPAMENTOS_ID);
    const aba = planilha.getActiveSheet();
    const dados = aba.getDataRange().getValues();
    
    if (dados.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const cabecalho = dados[0];
    const equipamentos = [];
    
    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const equipamento = {};
      
      for (let j = 0; j < cabecalho.length; j++) {
        equipamento[cabecalho[j]] = linha[j];
      }
      
      equipamentos.push(equipamento);
    }
    
    return ContentService.createTextOutput(JSON.stringify(equipamentos))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error('Erro no doGet:', error);
    return ContentService.createTextOutput(JSON.stringify({
      error: 'Erro ao carregar dados: ' + error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Processar login
function processarLogin(dados) {
  try {
    const planilhaUsuarios = SpreadsheetApp.openById(PLANILHA_USUARIOS_ID);
    const abaUsuarios = planilhaUsuarios.getActiveSheet();
    const dadosUsuarios = abaUsuarios.getDataRange().getValues();
    
    if (dadosUsuarios.length <= 1) {
      return { success: false, error: 'Nenhum usuário cadastrado' };
    }
    
    const cabecalho = dadosUsuarios[0];
    const emailIndex = cabecalho.indexOf('EMAIL');
    const senhaIndex = cabecalho.indexOf('SENHA');
    const nomeIndex = cabecalho.indexOf('NOME');
    const tipoIndex = cabecalho.indexOf('TIPO');
    const autorizadoIndex = cabecalho.indexOf('AUTORIZADO');
    
    if (emailIndex === -1 || senhaIndex === -1) {
      return { success: false, error: 'Estrutura da planilha de usuários inválida' };
    }
    
    // Procurar usuário
    for (let i = 1; i < dadosUsuarios.length; i++) {
      const linha = dadosUsuarios[i];
      
      if (linha[emailIndex] && linha[emailIndex].toString().toLowerCase() === dados.email.toLowerCase()) {
        if (linha[senhaIndex] && linha[senhaIndex].toString() === dados.senha) {
          // Verificar se está autorizado
          if (autorizadoIndex !== -1 && linha[autorizadoIndex] !== true && linha[autorizadoIndex] !== 'TRUE') {
            return { success: false, error: 'Usuário não autorizado. Aguarde aprovação do administrador.' };
          }
          
          return {
            success: true,
            usuario: {
              email: linha[emailIndex],
              nome: linha[nomeIndex] || 'Usuário',
              tipo: linha[tipoIndex] || 'operador',
              autorizado: linha[autorizadoIndex] || false
            }
          };
        } else {
          return { success: false, error: 'Senha incorreta' };
        }
      }
    }
    
    return { success: false, error: 'Usuário não encontrado' };
  } catch (error) {
    console.error('Erro no login:', error);
    return { success: false, error: 'Erro interno: ' + error.message };
  }
}

// Processar cadastro
function processarCadastro(dados) {
  try {
    const planilhaUsuarios = SpreadsheetApp.openById(PLANILHA_USUARIOS_ID);
    const abaUsuarios = planilhaUsuarios.getActiveSheet();
    const dadosUsuarios = abaUsuarios.getDataRange().getValues();
    
    // Verificar se o cabeçalho existe
    if (dadosUsuarios.length === 0) {
      // Criar cabeçalho se não existir
      abaUsuarios.getRange(1, 1, 1, 7).setValues([['NOME', 'EMAIL', 'SENHA', 'TIPO', 'AUTORIZADO', 'DATA_CADASTRO', 'DATA_ULTIMO_LOGIN']]);
    }
    
    const cabecalho = abaUsuarios.getRange(1, 1, 1, abaUsuarios.getLastColumn()).getValues()[0];
    const emailIndex = cabecalho.indexOf('EMAIL');
    const nomeIndex = cabecalho.indexOf('NOME');
    
    // Verificar se o e-mail já existe
    if (dadosUsuarios.length > 1) {
      for (let i = 1; i < dadosUsuarios.length; i++) {
        if (dadosUsuarios[i][emailIndex] && dadosUsuarios[i][emailIndex].toString().toLowerCase() === dados.email.toLowerCase()) {
          return { success: false, error: 'E-mail já cadastrado' };
        }
      }
    }
    
    // Gerar senha aleatória
    const senha = gerarSenhaAleatoria();
    
    // Adicionar novo usuário
    const novaLinha = abaUsuarios.getLastRow() + 1;
    const novoUsuario = [
      dados.nomeCompleto,  // NOME
      dados.email,         // EMAIL
      senha,               // SENHA
      'operador',          // TIPO
      false,               // AUTORIZADO
      new Date(),          // DATA_CADASTRO
      ''                   // DATA_ULTIMO_LOGIN
    ];
    
    abaUsuarios.getRange(novaLinha, 1, 1, novoUsuario.length).setValues([novoUsuario]);
    
    // Enviar senha por e-mail
    try {
      enviarSenhaPorEmail(dados.email, dados.nomeCompleto, senha);
    } catch (emailError) {
      console.error('Erro ao enviar e-mail:', emailError);
      // Continuar mesmo se o e-mail falhar
    }
    
    return {
      success: true,
      message: 'Cadastro realizado com sucesso! Você receberá uma senha por e-mail.'
    };
  } catch (error) {
    console.error('Erro no cadastro:', error);
    return { success: false, error: 'Erro interno: ' + error.message };
  }
}

// Atualizar equipamento
function atualizarEquipamento(dados) {
  try {
    const planilha = SpreadsheetApp.openById(PLANILHA_EQUIPAMENTOS_ID);
    const aba = planilha.getActiveSheet();
    const dadosExistentes = aba.getDataRange().getValues();
    
    if (dadosExistentes.length <= 1) {
      return { success: false, error: 'Planilha vazia ou sem dados' };
    }
    
    const cabecalho = dadosExistentes[0];
    const tagIndex = cabecalho.indexOf('TAG');
    
    if (tagIndex === -1) {
      return { success: false, error: 'Coluna TAG não encontrada' };
    }
    
    // Encontrar a linha do equipamento
    let linhaEncontrada = -1;
    for (let i = 1; i < dadosExistentes.length; i++) {
      if (dadosExistentes[i][tagIndex] === dados.TAG) {
        linhaEncontrada = i + 1; // +1 porque getRange usa índice baseado em 1
        break;
      }
    }
    
    if (linhaEncontrada === -1) {
      return { success: false, error: 'Equipamento não encontrado' };
    }
    
    // Atualizar apenas os campos fornecidos
    const camposParaAtualizar = ['STATUS', 'MOTIVO', 'PTS', 'OS', 'RETORNO', 'CADEADO', 'OBSERVACOES', 'MODIFICADO_POR', 'DATA'];
    
    camposParaAtualizar.forEach(campo => {
      const colunaIndex = cabecalho.indexOf(campo);
      if (colunaIndex !== -1 && dados.hasOwnProperty(campo)) {
        let valor = dados[campo];
        
        // Converter data se necessário
        if (campo === 'DATA' && valor) {
          valor = new Date(valor);
        } else if (campo === 'RETORNO' && valor) {
          valor = new Date(valor);
        }
        
        aba.getRange(linhaEncontrada, colunaIndex + 1).setValue(valor);
      }
    });
    
    return { success: true, message: 'Equipamento atualizado com sucesso' };
  } catch (error) {
    console.error('Erro ao atualizar equipamento:', error);
    return { success: false, error: 'Erro interno: ' + error.message };
  }
}

// Substituir planilha completa via CSV
function substituirPlanilhaCompleta(dados) {
  try {
    const planilha = SpreadsheetApp.openById(PLANILHA_EQUIPAMENTOS_ID);
    const aba = planilha.getActiveSheet();
    
    // Limpar toda a planilha
    aba.clear();
    
    // Definir cabeçalhos
    const cabecalhos = ['TAG', 'STATUS', 'MOTIVO', 'PTS', 'OS', 'RETORNO', 'CADEADO', 'OBSERVACOES', 'MODIFICADO_POR', 'DATA'];
    aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    
    // Adicionar dados
    if (dados.data && dados.data.length > 0) {
      const linhasParaInserir = dados.data.map(item => {
        return cabecalhos.map(cabecalho => {
          let valor = item[cabecalho] || '';
          
          // Converter datas se necessário
          if ((cabecalho === 'DATA' || cabecalho === 'RETORNO') && valor) {
            try {
              valor = new Date(valor);
            } catch (e) {
              // Se não conseguir converter, manter como string
            }
          }
          
          return valor;
        });
      });
      
      aba.getRange(2, 1, linhasParaInserir.length, cabecalhos.length).setValues(linhasParaInserir);
    }
    
    return { 
      success: true, 
      message: `Planilha substituída com sucesso! ${dados.data.length} registros foram carregados.` 
    };
  } catch (error) {
    console.error('Erro ao substituir planilha:', error);
    return { success: false, error: 'Erro interno: ' + error.message };
  }
}

// Atualizar via Excel VBA
function atualizarViaExcel(dados) {
  // Verificar token de segurança se fornecido
  if (dados.token && dados.token !== 'SEU_TOKEN_SEGURO') {
    return { success: false, error: 'Token de segurança inválido' };
  }
  
  // Usar a mesma lógica de atualização, mas com tipo específico
  return atualizarEquipamento(dados);
}

// Gerar senha aleatória
function gerarSenhaAleatoria() {
  const caracteres = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let senha = '';
  for (let i = 0; i < 8; i++) {
    senha += caracteres.charAt(Math.floor(Math.random() * caracteres.length));
  }
  return senha;
}

// Enviar senha por e-mail
function enviarSenhaPorEmail(email, nome, senha) {
  const assunto = 'Sua senha de acesso - Quadro de Disponibilidade';
  const corpo = `
Olá ${nome},

Seu cadastro foi realizado com sucesso no Sistema de Monitoramento de Equipamentos.

Seus dados de acesso:
E-mail: ${email}
Senha: ${senha}

IMPORTANTE: Seu acesso ainda precisa ser autorizado por um administrador. Você receberá uma notificação quando isso acontecer.

Atenciosamente,
Equipe Pernambuco III
  `;
  
  MailApp.sendEmail(email, assunto, corpo);
}

// Função para gerar PDF diário (será implementada na próxima fase)
function gerarPDFDiario() {
  try {
    console.log('Iniciando geração de PDF diário...');
    
    // Obter alterações do dia
    const alteracoesDoDia = obterAlteracoesDoDia();
    
    if (alteracoesDoDia.length === 0) {
      console.log('Nenhuma alteração encontrada para hoje.');
      return { success: true, message: 'Nenhuma alteração encontrada para hoje.' };
    }
    
    // Gerar HTML estilizado
    const htmlContent = gerarHTMLRelatorio(alteracoesDoDia);
    
    // Converter HTML em PDF
    const pdfBlob = gerarPDFDeHTML(htmlContent);
    
    // Enviar por e-mail
    enviarRelatorioPorEmail(pdfBlob, alteracoesDoDia.length);
    
    console.log(`PDF diário gerado e enviado com sucesso. ${alteracoesDoDia.length} alterações processadas.`);
    return { success: true, message: `PDF diário gerado e enviado com sucesso. ${alteracoesDoDia.length} alterações processadas.` };
  } catch (error) {
    console.error('Erro ao gerar PDF diário:', error);
    return { success: false, error: 'Erro ao gerar PDF diário: ' + error.message };
  }
}

// Obter alterações do dia atual
function obterAlteracoesDoDia() {
  try {
    const planilha = SpreadsheetApp.openById(PLANILHA_EQUIPAMENTOS_ID);
    const aba = planilha.getActiveSheet();
    const dados = aba.getDataRange().getValues();
    
    if (dados.length <= 1) {
      return [];
    }
    
    const cabecalho = dados[0];
    const dataIndex = cabecalho.indexOf('DATA');
    
    if (dataIndex === -1) {
      console.log('Coluna DATA não encontrada');
      return [];
    }
    
    const hoje = new Date();
    const inicioHoje = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
    const fimHoje = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate(), 23, 59, 59);
    
    const alteracoesDoDia = [];
    
    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const dataAlteracao = linha[dataIndex];
      
      if (dataAlteracao && dataAlteracao instanceof Date) {
        if (dataAlteracao >= inicioHoje && dataAlteracao <= fimHoje) {
          const equipamento = {};
          for (let j = 0; j < cabecalho.length; j++) {
            equipamento[cabecalho[j]] = linha[j];
          }
          alteracoesDoDia.push(equipamento);
        }
      }
    }
    
    return alteracoesDoDia;
  } catch (error) {
    console.error('Erro ao obter alterações do dia:', error);
    return [];
  }
}

// Gerar HTML estilizado para o relatório
function gerarHTMLRelatorio(alteracoes) {
  const hoje = new Date();
  const dataFormatada = hoje.toLocaleDateString('pt-BR', { 
    weekday: 'long', 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  });
  
  let html = `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Relatório Diário - ${dataFormatada}</title>
  <style>
    body {
      font-family: 'Arial', sans-serif;
      margin: 0;
      padding: 20px;
      background-color: #f8fafc;
      color: #334155;
      line-height: 1.6;
    }
    .header {
      text-align: center;
      margin-bottom: 30px;
      padding: 20px;
      background: linear-gradient(135deg, #2563eb 0%, #3b82f6 100%);
      color: white;
      border-radius: 10px;
    }
    .header h1 {
      margin: 0;
      font-size: 28px;
      font-weight: bold;
    }
    .header p {
      margin: 10px 0 0 0;
      font-size: 16px;
      opacity: 0.9;
    }
    .summary {
      background: white;
      padding: 20px;
      border-radius: 10px;
      margin-bottom: 30px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .summary h2 {
      margin: 0 0 15px 0;
      color: #1e293b;
      font-size: 20px;
    }
    .stats {
      display: flex;
      justify-content: space-around;
      text-align: center;
    }
    .stat-item {
      flex: 1;
    }
    .stat-number {
      font-size: 24px;
      font-weight: bold;
      color: #2563eb;
    }
    .stat-label {
      font-size: 14px;
      color: #64748b;
      margin-top: 5px;
    }
    .alteracoes {
      background: white;
      border-radius: 10px;
      overflow: hidden;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .alteracoes h2 {
      margin: 0;
      padding: 20px;
      background: #f1f5f9;
      color: #1e293b;
      font-size: 20px;
      border-bottom: 1px solid #e2e8f0;
    }
    .alteracao-item {
      padding: 20px;
      border-bottom: 1px solid #f1f5f9;
    }
    .alteracao-item:last-child {
      border-bottom: none;
    }
    .alteracao-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 15px;
    }
    .tag {
      font-size: 18px;
      font-weight: bold;
      color: #1e293b;
    }
    .status {
      padding: 4px 12px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: bold;
      text-transform: uppercase;
    }
    .status.ope {
      background: #dcfce7;
      color: #166534;
    }
    .status.st-by {
      background: #fef3c7;
      color: #92400e;
    }
    .status.manu {
      background: #fee2e2;
      color: #991b1b;
    }
    .alteracao-details {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 15px;
    }
    .detail-item {
      display: flex;
      flex-direction: column;
    }
    .detail-label {
      font-size: 12px;
      color: #64748b;
      font-weight: 500;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .detail-value {
      font-size: 14px;
      color: #1e293b;
      margin-top: 2px;
    }
    .footer {
      text-align: center;
      margin-top: 30px;
      padding: 20px;
      color: #64748b;
      font-size: 12px;
    }
    @media print {
      body { background-color: white; }
      .header, .summary, .alteracoes { box-shadow: none; }
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>Relatório Diário de Alterações</h1>
    <p>Quadro de Disponibilidade - Pernambuco III</p>
    <p>${dataFormatada}</p>
  </div>
  
  <div class="summary">
    <h2>Resumo do Dia</h2>
    <div class="stats">
      <div class="stat-item">
        <div class="stat-number">${alteracoes.length}</div>
        <div class="stat-label">Total de Alterações</div>
      </div>
      <div class="stat-item">
        <div class="stat-number">${alteracoes.filter(a => a.STATUS === 'OPE').length}</div>
        <div class="stat-label">Em Operação</div>
      </div>
      <div class="stat-item">
        <div class="stat-number">${alteracoes.filter(a => a.STATUS === 'ST-BY').length}</div>
        <div class="stat-label">Stand-by</div>
      </div>
      <div class="stat-item">
        <div class="stat-number">${alteracoes.filter(a => a.STATUS === 'MANU').length}</div>
        <div class="stat-label">Em Manutenção</div>
      </div>
    </div>
  </div>
  
  <div class="alteracoes">
    <h2>Detalhes das Alterações</h2>`;
  
  alteracoes.forEach(alteracao => {
    const statusClass = (alteracao.STATUS || '').toLowerCase().replace('-', '');
    const dataAlteracao = alteracao.DATA ? new Date(alteracao.DATA).toLocaleString('pt-BR') : 'N/A';
    const retorno = alteracao.RETORNO ? new Date(alteracao.RETORNO).toLocaleString('pt-BR') : '';
    
    html += `
    <div class="alteracao-item">
      <div class="alteracao-header">
        <div class="tag">${alteracao.TAG || 'N/A'}</div>
        <div class="status ${statusClass}">${alteracao.STATUS || 'N/A'}</div>
      </div>
      <div class="alteracao-details">`;
    
    if (alteracao.MOTIVO) {
      html += `
        <div class="detail-item">
          <div class="detail-label">Motivo</div>
          <div class="detail-value">${alteracao.MOTIVO}</div>
        </div>`;
    }
    
    if (alteracao.PTS) {
      html += `
        <div class="detail-item">
          <div class="detail-label">PTS</div>
          <div class="detail-value">${alteracao.PTS}</div>
        </div>`;
    }
    
    if (alteracao.OS) {
      html += `
        <div class="detail-item">
          <div class="detail-label">OS</div>
          <div class="detail-value">${alteracao.OS}</div>
        </div>`;
    }
    
    if (retorno) {
      html += `
        <div class="detail-item">
          <div class="detail-label">Previsão de Retorno</div>
          <div class="detail-value">${retorno}</div>
        </div>`;
    }
    
    if (alteracao.CADEADO) {
      html += `
        <div class="detail-item">
          <div class="detail-label">Cadeado</div>
          <div class="detail-value">${alteracao.CADEADO}</div>
        </div>`;
    }
    
    if (alteracao.OBSERVACOES) {
      html += `
        <div class="detail-item">
          <div class="detail-label">Observações</div>
          <div class="detail-value">${alteracao.OBSERVACOES}</div>
        </div>`;
    }
    
    html += `
        <div class="detail-item">
          <div class="detail-label">Modificado por</div>
          <div class="detail-value">${alteracao.MODIFICADO_POR || 'N/A'}</div>
        </div>
        <div class="detail-item">
          <div class="detail-label">Data/Hora</div>
          <div class="detail-value">${dataAlteracao}</div>
        </div>
      </div>
    </div>`;
  });
  
  html += `
  </div>
  
  <div class="footer">
    <p>Relatório gerado automaticamente em ${new Date().toLocaleString('pt-BR')}</p>
    <p>Sistema de Monitoramento de Equipamentos - Pernambuco III</p>
  </div>
</body>
</html>`;
  
  return html;
}

// Gerar PDF a partir do HTML
function gerarPDFDeHTML(htmlContent) {
  try {
    const blob = Utilities.newBlob(htmlContent, 'text/html', 'relatorio.html');
    
    // Converter HTML para PDF usando o serviço do Google
    const pdfBlob = Utilities.newBlob(
      DriveApp.createFile(blob).getBlob().getBytes(),
      'application/pdf',
      `relatorio_diario_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}.pdf`
    );
    
    // Remover arquivo HTML temporário
    const files = DriveApp.getFilesByName('relatorio.html');
    while (files.hasNext()) {
      DriveApp.moveFileToTrash(files.next());
    }
    
    return pdfBlob;
  } catch (error) {
    console.error('Erro ao gerar PDF:', error);
    throw new Error('Erro ao gerar PDF: ' + error.message);
  }
}

// Enviar relatório por e-mail
function enviarRelatorioPorEmail(pdfBlob, totalAlteracoes) {
  try {
    const hoje = new Date();
    const dataFormatada = hoje.toLocaleDateString('pt-BR');
    
    const assunto = `Relatório Diário - ${dataFormatada} - ${totalAlteracoes} alterações`;
    
    const corpo = `
Prezados,

Segue em anexo o relatório diário de alterações do Sistema de Monitoramento de Equipamentos.

Resumo do dia ${dataFormatada}:
• Total de alterações: ${totalAlteracoes}

O relatório completo está disponível no arquivo PDF em anexo.

Atenciosamente,
Sistema Automático de Relatórios
Pernambuco III
    `;
    
    // Lista de e-mails para envio (configurar conforme necessário)
    const destinatarios = [
      'admin@pernambuco3.com.br',
      'operacao@pernambuco3.com.br'
      // Adicionar mais e-mails conforme necessário
    ];
    
    destinatarios.forEach(email => {
      try {
        MailApp.sendEmail({
          to: email,
          subject: assunto,
          body: corpo,
          attachments: [pdfBlob]
        });
      } catch (emailError) {
        console.error(`Erro ao enviar e-mail para ${email}:`, emailError);
      }
    });
    
    console.log(`Relatório enviado para ${destinatarios.length} destinatários`);
  } catch (error) {
    console.error('Erro ao enviar relatório por e-mail:', error);
    throw new Error('Erro ao enviar relatório por e-mail: ' + error.message);
  }
}

// Função para configurar gatilhos automáticos
function configurarGatilhos() {
  // Deletar gatilhos existentes
  const gatilhos = ScriptApp.getProjectTriggers();
  gatilhos.forEach(gatilho => ScriptApp.deleteTrigger(gatilho));
  
  // Criar gatilho diário para PDF às 18:00
  ScriptApp.newTrigger('gerarPDFDiario')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();
}

