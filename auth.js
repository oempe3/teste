// auth.js - Sistema de Autenticação
// Endpoint do Google Apps Script
const AUTH_ENDPOINT = "https://script.google.com/macros/s/AKfycbwDEDoSIYOFaFf0I6I2XBwwNxd9tyDLoQTYPjttWX8iz8TpeY8ns1ddr4lFv-zuClYigg/exec";

// Verificar se o usuário está logado
function verificarLogin() {
  const usuarioLogado = sessionStorage.getItem('usuarioLogado');
  if (!usuarioLogado) {
    window.location.href = "index.html";
    return null;
  }
  
  try {
    const userData = JSON.parse(usuarioLogado);
    
    // Verificar se a sessão não expirou (24 horas)
    const loginTime = new Date(userData.loginTime);
    const agora = new Date();
    const diferencaHoras = (agora - loginTime) / (1000 * 60 * 60);
    
    if (diferencaHoras > 24) {
      logout();
      return null;
    }
    
    return userData;
  } catch (error) {
    console.error("Erro ao verificar login:", error);
    logout();
    return null;
  }
}

// Verificar se o usuário tem permissão de administrador
function verificarPermissaoAdmin() {
  const userData = verificarLogin();
  if (!userData || userData.tipo !== 'admin') {
    window.location.href = "status.html";
    return false;
  }
  return true;
}

// Fazer logout
function logout() {
  sessionStorage.removeItem('usuarioLogado');
  window.location.href = "index.html";
}

// Obter dados do usuário logado
function obterUsuarioLogado() {
  const usuarioLogado = sessionStorage.getItem('usuarioLogado');
  if (usuarioLogado) {
    try {
      return JSON.parse(usuarioLogado);
    } catch (error) {
      console.error("Erro ao obter dados do usuário:", error);
      return null;
    }
  }
  return null;
}

// Atualizar informações do usuário no header
function atualizarHeaderUsuario() {
  const userData = obterUsuarioLogado();
  if (userData) {
    const userInfoElement = document.getElementById('user-info');
    if (userInfoElement) {
      userInfoElement.textContent = `Olá, ${userData.nome}`;
    }
  }
}

// Proteger página (deve ser chamada no início de cada página protegida)
function protegerPagina(requerAdmin = false) {
  if (requerAdmin) {
    return verificarPermissaoAdmin();
  } else {
    return verificarLogin() !== null;
  }
}

