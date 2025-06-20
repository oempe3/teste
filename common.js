// common.js - Funções comuns do sistema
// Endpoint do Google Apps Script
const ENDPOINT = "https://script.google.com/macros/s/AKfycbwDEDoSIYOFaFf0I6I2XBwwNxd9tyDLoQTYPjttWX8iz8TpeY8ns1ddr4lFv-zuClYigg/exec";

// Atualizar relógio com horário de Brasília
function atualizarRelogio() {
  const agora = new Date();
  // Converter para horário de Brasília (UTC-3)
  const brasilia = new Date(agora.toLocaleString("en-US", {timeZone: "America/Sao_Paulo"}));
  const relogio = document.getElementById("relogio");
  if (relogio) {
    relogio.textContent = brasilia.toLocaleString("pt-BR", {
      weekday: 'short',
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    });
  }
}

// Função para mostrar mensagens
function mostrarMensagem(texto, tipo, elementoId = "mensagem") {
  const mensagem = document.getElementById(elementoId);
  if (mensagem) {
    mensagem.textContent = texto;
    mensagem.className = tipo === "success" ? "success-message" : "error-message show";
    mensagem.style.display = "block";
    
    if (tipo === "success") {
      mensagem.style.opacity = "1";
      mensagem.style.transform = "translateY(0)";
    }
    
    // Auto-hide após 5 segundos
    setTimeout(() => {
      mensagem.style.display = "none";
    }, 5000);
  }
}

// Função para fazer requisições HTTP
async function fazerRequisicao(url, dados = null, metodo = 'GET') {
  try {
    const opcoes = {
      method: metodo,
      headers: {
        "Content-Type": "application/json",
      }
    };
    
    if (dados && metodo !== 'GET') {
      opcoes.body = JSON.stringify(dados);
    }
    
    const response = await fetch(url, opcoes);
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    return await response.json();
  } catch (error) {
    console.error('Erro na requisição:', error);
    throw error;
  }
}

// Função para formatar data para o horário de Brasília
function formatarDataBrasilia(data, opcoes = {}) {
  if (!data) return 'N/A';
  
  const dataObj = typeof data === 'string' ? new Date(data) : data;
  
  const opcoesDefault = {
    timeZone: "America/Sao_Paulo",
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    ...opcoes
  };
  
  return dataObj.toLocaleString("pt-BR", opcoesDefault);
}

// Função para validar e-mail
function validarEmail(email) {
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}

// Função para escapar HTML
function escaparHTML(texto) {
  const div = document.createElement('div');
  div.textContent = texto;
  return div.innerHTML;
}

// Função para debounce (evitar muitas chamadas seguidas)
function debounce(func, wait) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

// Função para mostrar/esconder loading
function toggleLoading(buttonElement, show = true) {
  if (!buttonElement) return;
  
  const buttonText = buttonElement.querySelector('.button-text') || buttonElement.querySelector('span');
  const buttonLoader = buttonElement.querySelector('.button-loader') || buttonElement.querySelector('.loading-spinner');
  
  if (show) {
    if (buttonText) buttonText.style.display = 'none';
    if (buttonLoader) buttonLoader.style.display = 'inline-block';
    buttonElement.disabled = true;
  } else {
    if (buttonText) buttonText.style.display = 'inline';
    if (buttonLoader) buttonLoader.style.display = 'none';
    buttonElement.disabled = false;
  }
}

// Inicialização comum para todas as páginas
document.addEventListener('DOMContentLoaded', function() {
  // Iniciar relógio se existir elemento
  if (document.getElementById('relogio')) {
    atualizarRelogio();
    setInterval(atualizarRelogio, 1000);
  }
  
  // Adicionar classe fade-in aos elementos principais
  const mainContainer = document.querySelector('.main-container');
  if (mainContainer) {
    mainContainer.classList.add('fade-in');
  }
});

