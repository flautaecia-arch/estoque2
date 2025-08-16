// Configuração da API
const API_BASE = '/api';

// Estado da aplicação
let produtoAtual = null;

// Inicialização
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

function initializeApp() {
    setupTabNavigation();
    setupEventListeners();
    loadProdutos();
    loadContagens();
    loadResumo();
}

// Navegação por abas
function setupTabNavigation() {
    const tabButtons = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');

    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            const targetTab = button.getAttribute('data-tab');
            
            // Remove active class from all buttons and contents
            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabContents.forEach(content => content.classList.remove('active'));
            
            // Add active class to clicked button and corresponding content
            button.classList.add('active');
            document.getElementById(targetTab).classList.add('active');
        });
    });
}

// Event listeners
function setupEventListeners() {
    // Formulário de produto
    document.getElementById('form-produto').addEventListener('submit', handleProdutoSubmit);
    
    // Busca de produto
    document.getElementById('btn-buscar').addEventListener('click', buscarProduto);
    document.getElementById('busca-codigo').addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            buscarProduto();
        }
    });
    
    // Formulário de contagem
    document.getElementById('form-contagem').addEventListener('submit', handleContagemSubmit);
    
    // Botões de relatório
    document.getElementById('btn-relatorio-pdf').addEventListener('click', gerarRelatorioPDF);
    document.getElementById('btn-relatorio-excel').addEventListener('click', gerarRelatorioExcel);
    
    // Atualizar resumo
    document.getElementById('btn-atualizar-resumo').addEventListener('click', loadResumo);
    
    // Importação de produtos
    document.getElementById('btn-baixar-template').addEventListener('click', baixarTemplate);
    document.getElementById('btn-importar-excel').addEventListener('click', importarProdutos);
    
    // Modal
    setupModal();
}

// Modal
function setupModal() {
    const modal = document.getElementById('modal');
    const closeBtn = document.querySelector('.close');
    
    closeBtn.addEventListener('click', () => {
        modal.style.display = 'none';
    });
    
    window.addEventListener('click', (e) => {
        if (e.target === modal) {
            modal.style.display = 'none';
        }
    });
}

function showModal(message, type = 'info') {
    const modal = document.getElementById('modal');
    const modalBody = document.getElementById('modal-body');
    
    const icon = type === 'success' ? 'fa-check-circle' : 
                 type === 'error' ? 'fa-exclamation-circle' : 
                 'fa-info-circle';
    
    const color = type === 'success' ? '#27ae60' : 
                  type === 'error' ? '#e74c3c' : 
                  '#3498db';
    
    modalBody.innerHTML = `
        <div style="text-align: center; color: ${color};">
            <i class="fas ${icon}" style="font-size: 3rem; margin-bottom: 15px;"></i>
            <p style="font-size: 1.1rem; margin: 0;">${message}</p>
        </div>
    `;
    
    modal.style.display = 'block';
}

function showLoading() {
    document.getElementById('loading-overlay').style.display = 'flex';
}

function hideLoading() {
    document.getElementById('loading-overlay').style.display = 'none';
}

// Funções de API
async function apiRequest(url, options = {}) {
    try {
        showLoading();
        const response = await fetch(url, {
            headers: {
                'Content-Type': 'application/json',
                ...options.headers
            },
            ...options
        });
        
        const data = await response.json();
        
        if (!response.ok) {
            throw new Error(data.erro || `Erro HTTP: ${response.status}`);
        }
        
        return data;
    } catch (error) {
        console.error('Erro na API:', error);
        throw error;
    } finally {
        hideLoading();
    }
}

// Produtos
async function handleProdutoSubmit(e) {
    e.preventDefault();
    
    const formData = new FormData(e.target);
    const produto = {
        codigo: formData.get('codigo').trim(),
        nome: formData.get('nome').trim()
    };
    
    if (!produto.codigo || !produto.nome) {
        showModal('Por favor, preencha todos os campos.', 'error');
        return;
    }
    
    try {
        await apiRequest(`${API_BASE}/produtos`, {
            method: 'POST',
            body: JSON.stringify(produto)
        });
        
        showModal('Produto cadastrado com sucesso!', 'success');
        e.target.reset();
        loadProdutos();
    } catch (error) {
        showModal(`Erro ao cadastrar produto: ${error.message}`, 'error');
    }
}

async function loadProdutos() {
    try {
        const produtos = await apiRequest(`${API_BASE}/produtos`);
        renderProdutos(produtos);
    } catch (error) {
        console.error('Erro ao carregar produtos:', error);
    }
}

function renderProdutos(produtos) {
    const tbody = document.querySelector('#tabela-produtos tbody');
    tbody.innerHTML = '';
    
    produtos.forEach(produto => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${produto.codigo}</td>
            <td>${produto.nome}</td>
            <td>
                <button class="btn btn-danger btn-small" onclick="deletarProduto(${produto.id})">
                    <i class="fas fa-trash"></i> Excluir
                </button>
            </td>
        `;
        tbody.appendChild(row);
    });
}

async function deletarProduto(id) {
    if (!confirm('Tem certeza que deseja excluir este produto? Todas as contagens relacionadas também serão excluídas.')) {
        return;
    }
    
    try {
        await apiRequest(`${API_BASE}/produtos/${id}`, {
            method: 'DELETE'
        });
        
        showModal('Produto excluído com sucesso!', 'success');
        loadProdutos();
        loadContagens();
        loadResumo();
    } catch (error) {
        showModal(`Erro ao excluir produto: ${error.message}`, 'error');
    }
}

// Busca de produto
async function buscarProduto() {
    const codigo = document.getElementById('busca-codigo').value.trim();
    
    if (!codigo) {
        showModal('Por favor, digite um código de produto.', 'error');
        return;
    }
    
    try {
        const produto = await apiRequest(`${API_BASE}/produtos/${codigo}`);
        
        produtoAtual = produto;
        document.getElementById('produto-nome').textContent = produto.nome;
        document.getElementById('produto-codigo').textContent = produto.codigo;
        document.getElementById('produto-codigo-hidden').value = produto.codigo;
        
        document.getElementById('produto-encontrado').style.display = 'block';
        document.getElementById('card-contagem').style.display = 'block';
        
        // Limpar formulário de contagem
        document.getElementById('form-contagem').reset();
        document.getElementById('produto-codigo-hidden').value = produto.codigo;
        
    } catch (error) {
        showModal(`Produto não encontrado: ${error.message}`, 'error');
        document.getElementById('produto-encontrado').style.display = 'none';
        document.getElementById('card-contagem').style.display = 'none';
        produtoAtual = null;
    }
}

// Contagem
async function handleContagemSubmit(e) {
    e.preventDefault();
    
    if (!produtoAtual) {
        showModal('Por favor, busque um produto primeiro.', 'error');
        return;
    }
    
    const formData = new FormData(e.target);
    const contagem = {
        produto_codigo: formData.get('produto_codigo') || document.getElementById('produto-codigo-hidden').value,
        lote: formData.get('lote').trim(),
        validade_mes: parseInt(formData.get('validade_mes')),
        validade_ano: parseInt(formData.get('validade_ano')),
        quantidade: parseInt(formData.get('quantidade'))
    };
    
    if (!contagem.lote || !contagem.validade_mes || !contagem.validade_ano || !contagem.quantidade) {
        showModal('Por favor, preencha todos os campos.', 'error');
        return;
    }
    
    try {
        const resultado = await apiRequest(`${API_BASE}/contagens`, {
            method: 'POST',
            body: JSON.stringify(contagem)
        });
        
        showModal(resultado.mensagem, 'success');
        
        // Limpar apenas os campos de contagem, manter produto selecionado
        document.getElementById('lote').value = '';
        document.getElementById('validade-mes').value = '';
        document.getElementById('validade-ano').value = '';
        document.getElementById('quantidade').value = '';
        
        loadContagens();
        loadResumo();
    } catch (error) {
        showModal(`Erro ao registrar contagem: ${error.message}`, 'error');
    }
}

async function loadContagens() {
    try {
        const contagens = await apiRequest(`${API_BASE}/contagens`);
        renderContagens(contagens);
    } catch (error) {
        console.error('Erro ao carregar contagens:', error);
    }
}

function renderContagens(contagens) {
    const tbody = document.querySelector('#tabela-contagens tbody');
    tbody.innerHTML = '';
    
    contagens.forEach(contagem => {
        const row = document.createElement('tr');
        const meses = ['', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                      'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
        
        row.innerHTML = `
            <td>${contagem.produto.codigo}</td>
            <td>${contagem.produto.nome}</td>
            <td>${contagem.lote}</td>
            <td>${meses[contagem.validade_mes]}/${contagem.validade_ano}</td>
            <td>${contagem.quantidade}</td>
            <td>
                <button class="btn btn-danger btn-small" onclick="deletarContagem(${contagem.id})">
                    <i class="fas fa-trash"></i> Excluir
                </button>
            </td>
        `;
        tbody.appendChild(row);
    });
}

async function deletarContagem(id) {
    if (!confirm('Tem certeza que deseja excluir esta contagem?')) {
        return;
    }
    
    try {
        await apiRequest(`${API_BASE}/contagens/${id}`, {
            method: 'DELETE'
        });
        
        showModal('Contagem excluída com sucesso!', 'success');
        loadContagens();
        loadResumo();
    } catch (error) {
        showModal(`Erro ao excluir contagem: ${error.message}`, 'error');
    }
}

// Resumo
async function loadResumo() {
    const loadingElement = document.getElementById('resumo-loading');
    
    try {
        loadingElement.style.display = 'block';
        const resumo = await apiRequest(`${API_BASE}/relatorio/resumo`);
        renderResumo(resumo);
    } catch (error) {
        console.error('Erro ao carregar resumo:', error);
    } finally {
        loadingElement.style.display = 'none';
    }
}

function renderResumo(resumo) {
    const tbody = document.querySelector('#tabela-resumo tbody');
    const totalGeral = document.getElementById('total-geral');
    
    tbody.innerHTML = '';
    
    resumo.produtos.forEach(produto => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${produto.codigo}</td>
            <td>${produto.nome}</td>
            <td>${produto.total_quantidade}</td>
        `;
        tbody.appendChild(row);
    });
    
    totalGeral.innerHTML = `<strong>${resumo.total_geral}</strong>`;
}

// Relatórios
async function gerarRelatorioPDF() {
    try {
        showLoading();
        const response = await fetch(`${API_BASE}/relatorio/pdf`);
        
        if (!response.ok) {
            throw new Error('Erro ao gerar relatório PDF');
        }
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `relatorio_estoque_${new Date().toISOString().split('T')[0]}.pdf`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        showModal('Relatório PDF gerado com sucesso!', 'success');
    } catch (error) {
        showModal(`Erro ao gerar relatório PDF: ${error.message}`, 'error');
    } finally {
        hideLoading();
    }
}

async function gerarRelatorioExcel() {
    try {
        showLoading();
        const response = await fetch(`${API_BASE}/relatorio/excel`);
        
        if (!response.ok) {
            throw new Error('Erro ao gerar relatório Excel');
        }
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `relatorio_estoque_${new Date().toISOString().split('T')[0]}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        showModal('Relatório Excel gerado com sucesso!', 'success');
    } catch (error) {
        showModal(`Erro ao gerar relatório Excel: ${error.message}`, 'error');
    } finally {
        hideLoading();
    }
}

