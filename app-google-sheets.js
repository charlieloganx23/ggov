// ==================== CONFIGURAÇÃO GOOGLE SHEETS API ====================
// INSTRUÇÕES DE CONFIGURAÇÃO:
// 1. Acesse: https://console.cloud.google.com/
// 2. Crie um novo projeto ou selecione existente
// 3. Ative a API Google Sheets API
// 4. Crie credenciais (API Key)
// 5. Copie sua API Key e cole abaixo
// 6. Copie o ID da sua planilha (da URL) e cole abaixo

const GOOGLE_SHEETS_CONFIG = {
    apiKey: 'SUA_API_KEY_AQUI', // Substituir pela sua API Key
    spreadsheetId: 'SEU_SPREADSHEET_ID_AQUI', // ID da planilha (da URL)
    
    // Ranges das abas na planilha (Estrutura Tradicional)
    ranges: {
        // Informações do projeto (linha 3, colunas A-F)
        infoRow: 'Processo 1!A3:F3',
        // Etapas (a partir da linha 7, até 50 linhas)
        etapas: 'Processo 1!A7:K50',
        // Tarefas (a partir da linha 17, até 100 linhas)
        tarefas: 'Processo 1!A17:I100'
    },
    
    // Intervalo de atualização automática (em milissegundos)
    autoRefreshInterval: 30000 // 30 segundos
};

let autoRefreshTimer = null;
let isGoogleSheetsConnected = false;

// ==================== INICIALIZAÇÃO COM GOOGLE SHEETS ====================
document.addEventListener('DOMContentLoaded', function() {
    updateTimestamp();
    setInterval(updateTimestamp, 1000);
    
    initTabs();
    
    // Verificar se há configuração do Google Sheets
    if (GOOGLE_SHEETS_CONFIG.apiKey !== 'SUA_API_KEY_AQUI' && 
        GOOGLE_SHEETS_CONFIG.spreadsheetId !== 'SEU_SPREADSHEET_ID_AQUI') {
        // Carregar Google API
        loadGoogleSheetsAPI();
    } else {
        // Usar dados locais
        console.warn('⚠️ Google Sheets não configurado. Usando dados locais.');
        loadLocalData();
    }
    
    // Adicionar botões de controle
    addGoogleSheetsControls();
});

// ==================== CARREGAR GOOGLE SHEETS API ====================
function loadGoogleSheetsAPI() {
    const script = document.createElement('script');
    script.src = 'https://apis.google.com/js/api.js';
    script.onload = initGoogleAPI;
    document.head.appendChild(script);
}

function initGoogleAPI() {
    gapi.load('client', async () => {
        try {
            await gapi.client.init({
                apiKey: GOOGLE_SHEETS_CONFIG.apiKey,
                discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4']
            });
            
            console.log('✅ Google Sheets API conectada com sucesso!');
            isGoogleSheetsConnected = true;
            updateConnectionStatus(true);
            
            // Carregar dados iniciais
            await loadDataFromGoogleSheets();
            
            // Iniciar atualização automática
            startAutoRefresh();
            
        } catch (error) {
            console.error('❌ Erro ao conectar com Google Sheets:', error);
            updateConnectionStatus(false);
            loadLocalData();
        }
    });
}

// ==================== CARREGAR DADOS DO GOOGLE SHEETS ====================
async function loadDataFromGoogleSheets() {
    try {
        showLoading(true);
        
        // Buscar informações do projeto (linha 3, colunas A-G)
        const infoResponse = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEETS_CONFIG.spreadsheetId,
            range: GOOGLE_SHEETS_CONFIG.ranges.infoRow
        });
        
        // Atualizar informações do projeto
        if (infoResponse.result.values && infoResponse.result.values.length > 0) {
            const infoRow = infoResponse.result.values[0];
            processoData.sei = infoRow[0] || processoData.sei;
            processoData.prioridade = infoRow[1] || processoData.prioridade;
            processoData.categoria = infoRow[2] || processoData.categoria;
            processoData.dataInicio = infoRow[3] || processoData.dataInicio;
            processoData.dataTermino = infoRow[4] || processoData.dataTermino;
            processoData.descricao = infoRow[5] || processoData.descricao;
        }
        
        // Buscar dados das etapas
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEETS_CONFIG.spreadsheetId,
            range: GOOGLE_SHEETS_CONFIG.ranges.etapas
        });
        
        const rows = response.result.values;
        
        if (rows && rows.length > 0) {
            // Processar dados das etapas (filtrar linhas vazias)
            processoData.etapas = rows
                .filter(row => row[0] && row[0].trim() !== '') // Ignorar linhas sem nome de etapa
                .map((row, index) => ({
                    nome: row[0] || '',
                    status: row[1] || 'Não iniciada',
                    responsavel: row[2] || '',
                    dataInicio: row[3] || '',
                    dataTermino: row[4] || '',
                    produtos: row[5] || '',
                    dependencias: row[6] || '-',
                    progresso: parseFloat(row[7]) || 0,
                    horasEstimadas: parseInt(row[8]) || 0,
                    horasReais: parseInt(row[9]) || 0,
                    peso: parseFloat(row[10]) || 0.15
                }));
            
            // Buscar dados das tarefas
            const tarefasResponse = await gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: GOOGLE_SHEETS_CONFIG.spreadsheetId,
                range: GOOGLE_SHEETS_CONFIG.ranges.tarefas
            });
            
            const tarefasRows = tarefasResponse.result.values;
            
            if (tarefasRows && tarefasRows.length > 0) {
                // Filtrar linhas vazias nas tarefas também
                processoData.tarefas = tarefasRows
                    .filter(row => row[1] && row[1].trim() !== '') // Ignorar linhas sem nome de tarefa
                    .map(row => ({
                        etapa: row[0] || '',
                        nome: row[1] || '',
                        status: row[2] || 'Não iniciada',
                        responsavel: row[3] || '',
                        prioridade: row[4] || 'Média',
                        prazo: row[5] || '',
                        progresso: parseFloat(row[6]) || 0,
                        horas: parseInt(row[7]) || 0
                    }));
            }
            
            // Renderizar dados atualizados
            updateProcessoInfo();
            renderEtapas();
            renderTarefas();
            calcularKPIs();
            updateCharts();
            
            console.log('✅ Dados carregados do Google Sheets:', new Date().toLocaleTimeString());
            showNotification('✅ Dados atualizados com sucesso!', 'success');
            
        } else {
            console.warn('⚠️ Nenhum dado encontrado na planilha');
            loadLocalData();
        }
        
        showLoading(false);
        
    } catch (error) {
        console.error('❌ Erro ao carregar dados do Google Sheets:', error);
        
        // Log completo do erro para debug
        if (error.result && error.result.error) {
            console.error('📋 Detalhes do erro da API:');
            console.error('   Código:', error.result.error.code);
            console.error('   Mensagem:', error.result.error.message);
            console.error('   Status:', error.result.error.status);
            if (error.result.error.details) {
                console.error('   Detalhes:', error.result.error.details);
            }
        }
        
        // Mensagem de erro mais detalhada
        let errorMsg = 'Erro ao carregar dados da planilha';
        if (error.result && error.result.error) {
            const apiError = error.result.error;
            if (apiError.code === 400) {
                errorMsg = `Erro 400: ${apiError.message || 'Verifique o nome da aba e estrutura da planilha'}`;
            } else if (apiError.code === 403) {
                errorMsg = 'Planilha sem permissão de acesso (torne pública)';
            } else if (apiError.code === 404) {
                errorMsg = 'Planilha não encontrada - verifique o Spreadsheet ID';
            } else {
                errorMsg = `Erro ${apiError.code}: ${apiError.message}`;
            }
        }
        
        showNotification('❌ ' + errorMsg, 'error');
        showLoading(false);
        loadLocalData();
    }
}

// ==================== SALVAR DADOS NO GOOGLE SHEETS ====================
async function saveDataToGoogleSheets(etapaIndex, campo, valor) {
    if (!isGoogleSheetsConnected) {
        showNotification('⚠️ Google Sheets não conectado', 'warning');
        return;
    }
    
    try {
        // Mapear campo para coluna
        const colunaMap = {
            'status': 'B',
            'progresso': 'H',
            'horasReais': 'J'
        };
        
        const coluna = colunaMap[campo];
        if (!coluna) return;
        
        const linha = 12 + etapaIndex; // Linha inicial das etapas é 12
        const range = `Processo 1!${coluna}${linha}`;
        
        await gapi.client.sheets.spreadsheets.values.update({
            spreadsheetId: GOOGLE_SHEETS_CONFIG.spreadsheetId,
            range: range,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [[valor]]
            }
        });
        
        console.log(`✅ Atualizado: ${campo} = ${valor} na linha ${linha}`);
        showNotification('✅ Dados salvos na planilha!', 'success');
        
        // Recarregar dados após salvar
        setTimeout(() => loadDataFromGoogleSheets(), 1000);
        
    } catch (error) {
        console.error('❌ Erro ao salvar no Google Sheets:', error);
        showNotification('❌ Erro ao salvar dados', 'error');
    }
}

// ==================== AUTO REFRESH ====================
function startAutoRefresh() {
    if (autoRefreshTimer) {
        clearInterval(autoRefreshTimer);
    }
    
    autoRefreshTimer = setInterval(() => {
        if (isGoogleSheetsConnected) {
            console.log('🔄 Atualizando dados automaticamente...');
            loadDataFromGoogleSheets();
        }
    }, GOOGLE_SHEETS_CONFIG.autoRefreshInterval);
    
    console.log(`🔄 Auto-refresh ativado (a cada ${GOOGLE_SHEETS_CONFIG.autoRefreshInterval/1000}s)`);
}

function stopAutoRefresh() {
    if (autoRefreshTimer) {
        clearInterval(autoRefreshTimer);
        autoRefreshTimer = null;
        console.log('⏸️ Auto-refresh pausado');
    }
}

// ==================== CONTROLES DE INTERFACE ====================
function addGoogleSheetsControls() {
    const header = document.querySelector('.header .container');
    
    const controlsDiv = document.createElement('div');
    controlsDiv.className = 'google-controls';
    controlsDiv.innerHTML = `
        <div class="sync-controls">
            <button id="btn-refresh" class="btn-sync" onclick="manualRefresh()" title="Atualizar dados">
                <i class="fas fa-sync-alt"></i> Atualizar
            </button>
            <button id="btn-toggle-auto" class="btn-sync" onclick="toggleAutoRefresh()" title="Ativar/Desativar atualização automática">
                <i class="fas fa-clock"></i> Auto (30s)
            </button>
            <span id="connection-status" class="status-badge">
                <i class="fas fa-circle"></i> Desconectado
            </span>
            <button class="btn-sync" onclick="showConfigModal()" title="Configurar Google Sheets">
                <i class="fas fa-cog"></i>
            </button>
        </div>
    `;
    
    header.appendChild(controlsDiv);
}

function updateConnectionStatus(connected) {
    const statusBadge = document.getElementById('connection-status');
    if (statusBadge) {
        if (connected) {
            statusBadge.innerHTML = '<i class="fas fa-circle" style="color: #70AD47"></i> Conectado';
            statusBadge.style.background = '#C6EFCE';
            statusBadge.style.color = '#006100';
        } else {
            statusBadge.innerHTML = '<i class="fas fa-circle" style="color: #C00000"></i> Desconectado';
            statusBadge.style.background = '#FFC7CE';
            statusBadge.style.color = '#9C0006';
        }
    }
}

function manualRefresh() {
    const btn = document.getElementById('btn-refresh');
    btn.innerHTML = '<i class="fas fa-sync-alt fa-spin"></i> Atualizando...';
    btn.disabled = true;
    
    if (isGoogleSheetsConnected) {
        loadDataFromGoogleSheets().finally(() => {
            btn.innerHTML = '<i class="fas fa-sync-alt"></i> Atualizar';
            btn.disabled = false;
        });
    } else {
        showNotification('⚠️ Google Sheets não está conectado', 'warning');
        btn.innerHTML = '<i class="fas fa-sync-alt"></i> Atualizar';
        btn.disabled = false;
    }
}

function toggleAutoRefresh() {
    const btn = document.getElementById('btn-toggle-auto');
    
    if (autoRefreshTimer) {
        stopAutoRefresh();
        btn.style.opacity = '0.5';
        showNotification('⏸️ Atualização automática pausada', 'info');
    } else {
        startAutoRefresh();
        btn.style.opacity = '1';
        showNotification('▶️ Atualização automática ativada', 'success');
    }
}

function showConfigModal() {
    const modal = document.createElement('div');
    modal.className = 'config-modal';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h2><i class="fas fa-cog"></i> Configuração do Google Sheets</h2>
                <button onclick="closeConfigModal()" class="btn-close">&times;</button>
            </div>
            <div class="modal-body">
                <h3>📋 Passo a Passo:</h3>
                <ol class="config-steps">
                    <li>Acesse <a href="https://console.cloud.google.com/" target="_blank">Google Cloud Console</a></li>
                    <li>Crie um novo projeto ou selecione existente</li>
                    <li>No menu, vá em "APIs e Serviços" → "Biblioteca"</li>
                    <li>Procure e ative "Google Sheets API"</li>
                    <li>Vá em "Credenciais" → "Criar Credenciais" → "Chave de API"</li>
                    <li>Copie a API Key gerada</li>
                    <li>Abra sua planilha no Google Sheets</li>
                    <li>Copie o ID da planilha (da URL)</li>
                    <li>Cole as informações abaixo:</li>
                </ol>
                
                <div class="config-form">
                    <label>
                        <strong>API Key:</strong>
                        <input type="text" id="input-apikey" value="${GOOGLE_SHEETS_CONFIG.apiKey}" 
                               placeholder="AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX">
                    </label>
                    
                    <label>
                        <strong>Spreadsheet ID:</strong>
                        <input type="text" id="input-spreadsheet" value="${GOOGLE_SHEETS_CONFIG.spreadsheetId}" 
                               placeholder="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms">
                    </label>
                    
                    <div class="config-hint">
                        💡 <strong>Onde encontrar o Spreadsheet ID?</strong><br>
                        Na URL da planilha: docs.google.com/spreadsheets/d/<span style="background: yellow; color: black;">SPREADSHEET_ID</span>/edit
                    </div>
                    
                    <button onclick="saveConfig()" class="btn-save">
                        <i class="fas fa-save"></i> Salvar e Conectar
                    </button>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
}

function closeConfigModal() {
    const modal = document.querySelector('.config-modal');
    if (modal) modal.remove();
}

function saveConfig() {
    const apiKey = document.getElementById('input-apikey').value.trim();
    const spreadsheetId = document.getElementById('input-spreadsheet').value.trim();
    
    if (!apiKey || !spreadsheetId) {
        showNotification('⚠️ Preencha todos os campos', 'warning');
        return;
    }
    
    // Atualizar configuração
    GOOGLE_SHEETS_CONFIG.apiKey = apiKey;
    GOOGLE_SHEETS_CONFIG.spreadsheetId = spreadsheetId;
    
    // Salvar no localStorage
    localStorage.setItem('googleSheetsConfig', JSON.stringify({
        apiKey: apiKey,
        spreadsheetId: spreadsheetId
    }));
    
    closeConfigModal();
    showNotification('✅ Configuração salva! Reconectando...', 'success');
    
    // Reconectar
    setTimeout(() => {
        loadGoogleSheetsAPI();
    }, 1000);
}

// ==================== CARREGAR CONFIGURAÇÃO SALVA ====================
function loadSavedConfig() {
    const saved = localStorage.getItem('googleSheetsConfig');
    if (saved) {
        try {
            const config = JSON.parse(saved);
            GOOGLE_SHEETS_CONFIG.apiKey = config.apiKey;
            GOOGLE_SHEETS_CONFIG.spreadsheetId = config.spreadsheetId;
            console.log('✅ Configuração carregada do localStorage');
        } catch (e) {
            console.error('❌ Erro ao carregar configuração salva');
        }
    }
}

// Carregar configuração ao iniciar
loadSavedConfig();

// ==================== NOTIFICAÇÕES ====================
function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.textContent = message;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.classList.add('show');
    }, 100);
    
    setTimeout(() => {
        notification.classList.remove('show');
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}

function showLoading(show) {
    let loader = document.getElementById('loading-overlay');
    
    if (show && !loader) {
        loader = document.createElement('div');
        loader.id = 'loading-overlay';
        loader.innerHTML = `
            <div class="spinner">
                <i class="fas fa-sync-alt fa-spin"></i>
                <p>Carregando dados...</p>
            </div>
        `;
        document.body.appendChild(loader);
    } else if (!show && loader) {
        loader.remove();
    }
}

// ==================== DADOS LOCAIS (FALLBACK) ====================
const processoDataLocal = {
    id: 1,
    nome: "Mapeamento dos processos do Gabinete de Governança",
    sei: "0000000000000",
    prioridade: "Alta",
    categoria: "Mapeamento",
    descricao: "Realizar o mapeamento completo dos processos administrativos e operacionais do Gabinete de Governança (GGOV)...",
    dataInicio: "2025-12-10",
    dataTermino: "2026-01-31",
    orcamento: 15000,
    etapas: [
        {
            nome: "Levantamento de Informações",
            status: "Em execução",
            responsavel: "Luma Damon de Oliveira Melo",
            dataInicio: "2025-12-10",
            dataTermino: "2026-01-16",
            produtos: "Plano do projeto",
            dependencias: "-",
            progresso: 0.70,
            horasEstimadas: 80,
            horasReais: 56,
            peso: 0.15
        },
        {
            nome: "Mapeamento de Processos",
            status: "Em execução",
            responsavel: "Suerlei Gondim Dutra",
            dataInicio: "2025-12-10",
            dataTermino: "2026-01-31",
            produtos: "Relatório de Levantamento\nMapas de Processos\nRelatório de Análise",
            dependencias: "Etapa 1",
            progresso: 0.60,
            horasEstimadas: 120,
            horasReais: 72,
            peso: 0.25
        },
        {
            nome: "Análise de Processos",
            status: "Não iniciada",
            responsavel: "Equipe GGOV",
            dataInicio: "2026-01-17",
            dataTermino: "2026-02-15",
            produtos: "Análise de eficiência e gargalos",
            dependencias: "Etapa 1, 2",
            progresso: 0.00,
            horasEstimadas: 100,
            horasReais: 0,
            peso: 0.20
        },
        {
            nome: "Documentação e Relatório Final",
            status: "Não iniciada",
            responsavel: "Equipe Técnica",
            dataInicio: "2026-02-01",
            dataTermino: "2026-02-28",
            produtos: "Relatório Final Consolidado",
            dependencias: "Etapa 3",
            progresso: 0.00,
            horasEstimadas: 80,
            horasReais: 0,
            peso: 0.20
        },
        {
            nome: "Validação e Aprovação",
            status: "Não iniciada",
            responsavel: "Direção GGOV",
            dataInicio: "2026-02-20",
            dataTermino: "2026-03-10",
            produtos: "Aprovação formal",
            dependencias: "Etapa 4",
            progresso: 0.00,
            horasEstimadas: 40,
            horasReais: 0,
            peso: 0.10
        },
        {
            nome: "Entrega e Implementação",
            status: "Não iniciada",
            responsavel: "Equipe GGOV Completa",
            dataInicio: "2026-03-01",
            dataTermino: "2026-03-31",
            produtos: "Processos implementados e ativos",
            dependencias: "Etapa 5",
            progresso: 0.00,
            horasEstimadas: 60,
            horasReais: 0,
            peso: 0.10
        }
    ],
    tarefas: [
        {
            etapa: "Etapa 1",
            nome: "1. Realizar entrevistas com os responsáveis de cada área",
            status: "Em execução",
            responsavel: "Luma Damon",
            prioridade: "Alta",
            prazo: "2025-12-15",
            progresso: 0.80,
            horas: 20
        },
        {
            etapa: "Etapa 1",
            nome: "2. Analisar documentos existentes, como manuais e fluxos anteriores",
            status: "Em execução",
            responsavel: "Luma Damon",
            prioridade: "Alta",
            prazo: "2025-12-20",
            progresso: 0.70,
            horas: 16
        },
        {
            etapa: "Etapa 1",
            nome: "3. Observar e registrar as atividades e etapas realizadas nas áreas de governança",
            status: "Em execução",
            responsavel: "Luma Damon",
            prioridade: "Média",
            prazo: "2026-01-05",
            progresso: 0.60,
            horas: 24
        },
        {
            etapa: "Etapa 1",
            nome: "4. Criar um questionário para coletar dados com os responsáveis pelos processos",
            status: "Concluída",
            responsavel: "Luma Damon",
            prioridade: "Alta",
            prazo: "2025-12-12",
            progresso: 1.00,
            horas: 8
        },
        {
            etapa: "Etapa 1",
            nome: "5. Identificar as entradas, saídas, recursos e responsáveis de cada processo",
            status: "Em execução",
            responsavel: "Luma Damon",
            prioridade: "Alta",
            prazo: "2026-01-10",
            progresso: 0.50,
            horas: 12
        },
        {
            etapa: "Etapa 2",
            nome: "1. Documentar processos no formato AS-IS",
            status: "Em execução",
            responsavel: "Suerlei Gondim",
            prioridade: "Alta",
            prazo: "2026-01-15",
            progresso: 0.70,
            horas: 40
        },
        {
            etapa: "Etapa 2",
            nome: "2. Criar diagramas de fluxo (BPMN)",
            status: "Em execução",
            responsavel: "Suerlei Gondim",
            prioridade: "Alta",
            prazo: "2026-01-20",
            progresso: 0.60,
            horas: 30
        },
        {
            etapa: "Etapa 2",
            nome: "3. Identificar gargalos e ineficiências",
            status: "Não iniciada",
            responsavel: "Suerlei Gondim",
            prioridade: "Média",
            prazo: "2026-01-25",
            progresso: 0.00,
            horas: 25
        },
        {
            etapa: "Etapa 2",
            nome: "4. Consolidar relatório de levantamento",
            status: "Não iniciada",
            responsavel: "Suerlei Gondim",
            prioridade: "Média",
            prazo: "2026-01-31",
            progresso: 0.00,
            horas: 25
        }
    ]
};

let processoData = { ...processoDataLocal };

function loadLocalData() {
    processoData = { ...processoDataLocal };
    renderEtapas();
    renderTarefas();
    calcularKPIs();
    initCharts();
}

// ==================== RESTO DO CÓDIGO ORIGINAL ====================
// (manter todo o código original de timestamp, tabs, render, etc.)

function updateTimestamp() {
    const now = new Date();
    const formatted = now.toLocaleString('pt-BR', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    });
    const timestampEl = document.getElementById('timestamp');
    if (timestampEl) {
        timestampEl.textContent = formatted;
    }
}

function updateProcessoInfo() {
    // Atualizar SEI
    const seiIndicator = document.getElementById('sei-indicator');
    const seiValue = document.getElementById('sei-value');
    if (seiIndicator) seiIndicator.textContent = processoData.sei || '0000000000000';
    if (seiValue) seiValue.textContent = processoData.sei || '0000000000000';
    
    // Atualizar Prioridade
    const prioridadeBadge = document.getElementById('prioridade-badge');
    if (prioridadeBadge) {
        prioridadeBadge.textContent = processoData.prioridade || 'Alta';
        prioridadeBadge.className = 'badge badge-' + (processoData.prioridade || 'alta').toLowerCase();
    }
    
    // Atualizar Categoria
    const categoriaValue = document.getElementById('categoria-value');
    if (categoriaValue) categoriaValue.textContent = processoData.categoria || 'Mapeamento';
    
    // Atualizar Datas
    const dataInicioValue = document.getElementById('data-inicio-value');
    const dataTerminoValue = document.getElementById('data-termino-value');
    if (dataInicioValue) dataInicioValue.textContent = processoData.dataInicio || '10/12/2025';
    if (dataTerminoValue) dataTerminoValue.textContent = processoData.dataTermino || '31/01/2026';
}

function initTabs() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    
    tabBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            const tabId = this.dataset.tab;
            
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            
            this.classList.add('active');
            document.getElementById(tabId).classList.add('active');
        });
    });
}

function expandProcesso(id) {
    const detailsDiv = document.getElementById(`processo-${id}-details`);
    const btn = event.target.closest('.btn-expand');
    
    if (detailsDiv.style.display === 'none') {
        detailsDiv.style.display = 'block';
        btn.innerHTML = '<i class="fas fa-chevron-up"></i> Recolher';
    } else {
        detailsDiv.style.display = 'none';
        btn.innerHTML = '<i class="fas fa-chevron-down"></i> Expandir Detalhes';
    }
}

function renderEtapas() {
    const tbody = document.getElementById('etapas-tbody');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    processoData.etapas.forEach((etapa, index) => {
        const tr = document.createElement('tr');
        
        const statusClass = etapa.status.toLowerCase().replace(' ', '-');
        const progressPercent = Math.round(etapa.progresso * 100);
        
        tr.innerHTML = `
            <td><strong>${etapa.nome}</strong></td>
            <td><span class="status-badge status-${statusClass}">${etapa.status}</span></td>
            <td>${etapa.responsavel}</td>
            <td>${formatDate(etapa.dataInicio)}</td>
            <td>${formatDate(etapa.dataTermino)}</td>
            <td style="max-width: 250px; font-size: 0.85rem;">${etapa.produtos}</td>
            <td>
                <div class="progress-bar" style="height: 25px;">
                    <div class="progress-fill" style="width: ${progressPercent}%">${progressPercent}%</div>
                </div>
            </td>
            <td>${etapa.horasEstimadas}h / ${etapa.horasReais}h</td>
            <td>${Math.round(etapa.peso * 100)}%</td>
            <td>
                <button class="btn-expand" style="padding: 5px 10px; font-size: 0.85rem;" onclick="verDetalhesEtapa(${index})">
                    <i class="fas fa-eye"></i> Ver
                </button>
            </td>
        `;
        
        tbody.appendChild(tr);
    });
}

function renderTarefas() {
    const tbody = document.getElementById('tarefas-tbody');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    processoData.tarefas.forEach(tarefa => {
        const tr = document.createElement('tr');
        tr.dataset.etapa = tarefa.etapa;
        
        const statusClass = tarefa.status.toLowerCase().replace(' ', '-');
        const prioridadeClass = tarefa.prioridade.toLowerCase();
        const progressPercent = Math.round(tarefa.progresso * 100);
        
        tr.innerHTML = `
            <td><strong>${tarefa.etapa}</strong></td>
            <td style="max-width: 300px;">${tarefa.nome}</td>
            <td><span class="status-badge status-${statusClass}">${tarefa.status}</span></td>
            <td>${tarefa.responsavel}</td>
            <td><span class="badge badge-${prioridadeClass}">${tarefa.prioridade}</span></td>
            <td>${formatDate(tarefa.prazo)}</td>
            <td>
                <div class="progress-bar" style="height: 20px;">
                    <div class="progress-fill" style="width: ${progressPercent}%; font-size: 0.8rem;">${progressPercent}%</div>
                </div>
            </td>
            <td>${tarefa.horas}h</td>
        `;
        
        tbody.appendChild(tr);
    });
}

function filtrarTarefas() {
    const filtro = document.getElementById('filtro-etapa').value;
    const rows = document.querySelectorAll('#tarefas-tbody tr');
    
    rows.forEach(row => {
        if (filtro === 'all' || row.dataset.etapa === filtro) {
            row.style.display = '';
        } else {
            row.style.display = 'none';
        }
    });
}

let statusChart, progressChart;

function initCharts() {
    const statusCtx = document.getElementById('statusChart');
    const progressCtx = document.getElementById('progressChart');
    
    if (!statusCtx || !progressCtx) return;
    
    // Destruir gráficos existentes antes de criar novos
    if (statusChart) {
        statusChart.destroy();
        statusChart = null;
    }
    if (progressChart) {
        progressChart.destroy();
        progressChart = null;
    }
    
    const statusCount = {
        'Em execução': 0,
        'Concluída': 0,
        'Não iniciada': 0
    };
    
    processoData.etapas.forEach(etapa => {
        statusCount[etapa.status]++;
    });
    
    statusChart = new Chart(statusCtx.getContext('2d'), {
        type: 'doughnut',
        data: {
            labels: Object.keys(statusCount),
            datasets: [{
                data: Object.values(statusCount),
                backgroundColor: ['#FFA500', '#70AD47', '#5B9BD5']
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: { legend: { position: 'bottom' } }
        }
    });
    
    progressChart = new Chart(progressCtx.getContext('2d'), {
        type: 'bar',
        data: {
            labels: processoData.etapas.map(e => e.nome.substring(0, 20) + '...'),
            datasets: [{
                label: '% Progresso',
                data: processoData.etapas.map(e => e.progresso * 100),
                backgroundColor: '#4472C4'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: { y: { beginAtZero: true, max: 100 } },
            plugins: { legend: { display: false } }
        }
    });
}

function updateCharts() {
    if (statusChart && progressChart) {
        const statusCount = {
            'Em execução': 0,
            'Concluída': 0,
            'Não iniciada': 0
        };
        
        processoData.etapas.forEach(etapa => {
            statusCount[etapa.status]++;
        });
        
        statusChart.data.datasets[0].data = Object.values(statusCount);
        statusChart.update();
        
        progressChart.data.datasets[0].data = processoData.etapas.map(e => e.progresso * 100);
        progressChart.update();
    }
}

function calcularKPIs() {
    let somaProgresso = 0;
    processoData.etapas.forEach(etapa => {
        somaProgresso += etapa.progresso * etapa.peso;
    });
    
    const percentualGeral = Math.round(somaProgresso * 100);
    
    const emExecucao = processoData.etapas.filter(e => e.status === 'Em execução').length;
    const concluidas = processoData.etapas.filter(e => e.status === 'Concluída').length;
    const naoIniciadas = processoData.etapas.filter(e => e.status === 'Não iniciada').length;
    
    const kpiAtivos = document.getElementById('kpi-ativos');
    const kpiConcluidos = document.getElementById('kpi-concluidos');
    const kpiPlanejados = document.getElementById('kpi-planejados');
    const kpiProgresso = document.getElementById('kpi-progresso');
    
    if (kpiAtivos) kpiAtivos.textContent = emExecucao;
    if (kpiConcluidos) kpiConcluidos.textContent = concluidas;
    if (kpiPlanejados) kpiPlanejados.textContent = naoIniciadas;
    if (kpiProgresso) kpiProgresso.textContent = percentualGeral + '%';
    
    const alertsDiv = document.getElementById('alerts-container');
    if (alertsDiv) {
        alertsDiv.innerHTML = '';
        
        if (percentualGeral < 50) {
            alertsDiv.innerHTML += '<div class="alert alert-warning">🟡 ATENÇÃO: Projeto com menos de 50% de conclusão</div>';
        }
        
        if (emExecucao > 0) {
            alertsDiv.innerHTML += '<div class="alert alert-warning">🟡 ATENÇÃO: ' + emExecucao + ' etapa(s) em execução necessitam acompanhamento</div>';
        }
        
        if (percentualGeral >= 70) {
            alertsDiv.innerHTML += '<div class="alert alert-success">🟢 SUCESSO: Projeto com ótimo progresso!</div>';
        }
    }
}

function formatDate(dateString) {
    const date = new Date(dateString);
    return date.toLocaleDateString('pt-BR');
}

function verDetalhesEtapa(index) {
    alert('Funcionalidade em desenvolvimento!\nEtapa: ' + processoData.etapas[index].nome);
}

console.log('🚀 Sistema GGOV Revolucionário com Google Sheets carregado!');
