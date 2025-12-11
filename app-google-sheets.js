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
    
    // DETECÇÃO AUTOMÁTICA DE ABAS
    // O sistema detecta automaticamente qualquer aba que contenha este padrão no nome
    padraoNomeProcesso: 'Processo', // Abas que contêm "Processo" serão carregadas
    
    // Intervalo de atualização automática (em milissegundos)
    // autoRefreshInterval: 60000 // 1 minuto - DESATIVADO TEMPORARIAMENTE
    autoRefreshInterval: 0 // Auto-refresh desativado (use o botão "Atualizar" manualmente)
};

let autoRefreshTimer = null;
let isGoogleSheetsConnected = false;
let todosProcessos = []; // Array para armazenar todos os processos

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

// ==================== DETECTAR ABAS AUTOMATICAMENTE ====================
async function detectarAbasProcessos() {
    try {
        // Buscar metadados da planilha
        const response = await gapi.client.sheets.spreadsheets.get({
            spreadsheetId: GOOGLE_SHEETS_CONFIG.spreadsheetId
        });
        
        // Extrair nomes das abas
        const sheets = response.result.sheets || [];
        const nomesAbas = sheets.map(sheet => sheet.properties.title);
        
        // Filtrar abas que contêm o padrão configurado
        const abasProcessos = nomesAbas.filter(nome => 
            nome.includes(GOOGLE_SHEETS_CONFIG.padraoNomeProcesso)
        );
        
        console.log(`🔍 Abas detectadas (${abasProcessos.length}):`, abasProcessos);
        return abasProcessos;
        
    } catch (error) {
        console.error('❌ Erro ao detectar abas:', error);
        return [];
    }
}

// ==================== CARREGAR DADOS DO GOOGLE SHEETS ====================
async function loadDataFromGoogleSheets() {
    try {
        showLoading(true);
        todosProcessos = []; // Limpar array
        
        // Detectar abas automaticamente
        const abasEncontradas = await detectarAbasProcessos();
        
        if (abasEncontradas.length === 0) {
            console.warn('⚠️ Nenhuma aba de processo encontrada');
            showNotification('⚠️ Nenhum processo encontrado na planilha', 'warning');
            showLoading(false);
            loadLocalData();
            return;
        }
        
        // Carregar todos os processos detectados
        for (const nomeProcesso of abasEncontradas) {
            await loadProcessoData(nomeProcesso);
        }
        
        // Renderizar todos os processos
        renderTodosProcessos();
        calcularKPIsGlobais();
        updateCharts();
        
        console.log('✅ Dados de todos os processos carregados:', new Date().toLocaleTimeString());
        showNotification(`✅ ${todosProcessos.length} processo(s) atualizado(s)!`, 'success');
        showLoading(false);
        
    } catch (error) {
        console.error('❌ Erro ao carregar dados:', error);
        showNotification('❌ Erro ao carregar dados', 'error');
        showLoading(false);
        loadLocalData();
    }
}

async function loadProcessoData(nomeProcesso) {
    try {
        // Escapar aspas simples no nome da aba e envolver em aspas para caracteres especiais
        const nomeAbaEscapado = `'${nomeProcesso.replace(/'/g, "''")}'`;
        
        const ranges = {
            infoLinha2: `${nomeAbaEscapado}!A2:G2`,
            infoLinha3: `${nomeAbaEscapado}!A3:H3`,
            infoLinha4: `${nomeAbaEscapado}!A4:H4`,
            infoLinha5: `${nomeAbaEscapado}!A5:D5`,
            etapas: `${nomeAbaEscapado}!A8:N50`,
            tarefas: `${nomeAbaEscapado}!A17:I100`
        };
        
        const processo = {
            nome: nomeProcesso,
            // Campos originais
            sei: '',
            prioridade: '',
            categoria: '',
            dataInicio: '',
            dataTermino: '',
            descricao: '',
            // Campos novos da estrutura híbrida
            unidadeDemandante: '',
            responsavelDemanda: '',
            duracaoTotal: '',
            diasRestantes: '',
            progressoGeral: '',
            statusAtual: '',
            etapas: [],
            tarefas: []
        };
        
        // Buscar informações do projeto - Linha 2
        const info2Response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEETS_CONFIG.spreadsheetId,
            range: ranges.infoLinha2
        });
        if (info2Response.result.values && info2Response.result.values.length > 0) {
            const row = info2Response.result.values[0];
            // A2:G2 = Projeto/Demanda: | nome | Descrição: | desc | Unidade Demandante: | unidade
            processo.descricao = row[3] || '';
            processo.unidadeDemandante = row[5] || '';
        }
        
        // Buscar informações do projeto - Linha 3
        const info3Response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEETS_CONFIG.spreadsheetId,
            range: ranges.infoLinha3
        });
        if (info3Response.result.values && info3Response.result.values.length > 0) {
            const row = info3Response.result.values[0];
            // A3:H3 = Número SEI: | SEI | Categoria: | cat | Prioridade: | pri | Responsável Demanda: | resp
            processo.sei = row[1] || '';
            processo.categoria = row[3] || '';
            processo.prioridade = row[5] || '';
            processo.responsavelDemanda = row[7] || '';
        }
        
        // Buscar informações do projeto - Linha 4
        const info4Response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEETS_CONFIG.spreadsheetId,
            range: ranges.infoLinha4
        });
        if (info4Response.result.values && info4Response.result.values.length > 0) {
            const row = info4Response.result.values[0];
            // A4:H4 = Data Início: | data | Data Término: | data | Duração (dias): | dur | Dias Restantes: | dias
            // A4 = label, B4 = data inicio, C4 = label, D4 = data termino, E4 = label, F4 = duracao, G4 = label, H4 = dias restantes
            processo.dataInicio = row[1] || '';
            processo.dataTermino = row[3] || '';
            processo.duracaoTotal = row[5] || '';
            processo.diasRestantes = row[7] || ''; // Coluna H (índice 7)
        }
        
        // Buscar informações do projeto - Linha 5
        const info5Response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEETS_CONFIG.spreadsheetId,
            range: ranges.infoLinha5
        });
        if (info5Response.result.values && info5Response.result.values.length > 0) {
            const row = info5Response.result.values[0];
            // A5:C5 = Progresso Geral: | % | Status Atual: | status
            // A5 = label, B5 = progresso, C5 = label "Status Atual:", D5 = valor do status
            processo.progressoGeral = row[1] || '';
            processo.statusAtual = row[3] || ''; // Coluna D (índice 3)
        }
        
        // Buscar dados das etapas
        const etapasResponse = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEETS_CONFIG.spreadsheetId,
            range: ranges.etapas
        });
        
        if (etapasResponse.result.values && etapasResponse.result.values.length > 0) {
            processo.etapas = etapasResponse.result.values
                .filter(row => row[0] && row[0].trim() !== '')
                .map(row => ({
                    // Campos originais (A-K)
                    nome: row[0] || '',
                    status: row[1] || 'Não iniciada',
                    responsavel: row[2] || '',
                    dataInicio: row[3] || '',
                    dataTermino: row[4] || '',
                    produtos: row[5] || '',
                    dependencias: row[6] || '-',
                    progresso: (parseFloat(row[7]) || 0) / 100, // Converter % para decimal
                    horasEstimadas: parseInt(row[8]) || 0,
                    horasReais: parseInt(row[9]) || 0,
                    peso: (parseFloat(row[10]) || 15) / 100, // Converter % para decimal
                    // Campos novos híbridos (L-N)
                    situacao: row[11] || '',
                    tarefasTexto: row[12] || '',
                    observacoes: row[13] || ''
                }));
        }
        
        // Buscar dados das tarefas
        const tarefasResponse = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEETS_CONFIG.spreadsheetId,
            range: ranges.tarefas
        });
        
        if (tarefasResponse.result.values && tarefasResponse.result.values.length > 0) {
            processo.tarefas = tarefasResponse.result.values
                .filter(row => row[1] && row[1].trim() !== '')
                .map(row => ({
                    etapa: row[0] || '',
                    nome: row[1] || '',
                    status: row[2] || 'Não iniciada',
                    responsavel: row[3] || '',
                    prioridade: row[4] || 'Média',
                    prazo: row[5] || '',
                    progresso: (parseFloat(row[6]) || 0) / 100, // Converter % para decimal
                    horas: parseInt(row[7]) || 0
                }));
        }
        
        todosProcessos.push(processo);
        console.log(`✅ Processo carregado: ${nomeProcesso}`);
        
    } catch (error) {
        console.error(`❌ Erro ao carregar ${nomeProcesso}:`, error);
        throw error;
    }
}

// ==================== RENDERIZAR TODOS OS PROCESSOS ====================
function renderTodosProcessos() {
    const commandCenter = document.getElementById('command-center');
    if (!commandCenter) return;
    
    // Encontrar container de cards ou criar
    let cardsContainer = commandCenter.querySelector('.processos-container');
    if (!cardsContainer) {
        const title = commandCenter.querySelector('.section-title');
        cardsContainer = document.createElement('div');
        cardsContainer.className = 'processos-container';
        title.after(cardsContainer);
    }
    
    cardsContainer.innerHTML = '';
    
    todosProcessos.forEach((proc, index) => {
        const card = criarCardProcesso(proc, index);
        cardsContainer.appendChild(card);
    });
    
    // Criar abas de navegação e conteúdo para cada processo
    criarAbasProcessos();
    
    // Atualizar primeiro processo na aba detalhada (compatibilidade)
    if (todosProcessos.length > 0) {
        processoData = todosProcessos[0];
        updateProcessoInfo();
        renderEtapas();
        renderTarefas();
    }
}

function criarCardProcesso(proc, index) {
    const card = document.createElement('div');
    card.className = 'processo-card';
    card.setAttribute('data-processo', index + 1);
    
    // Calcular métricas
    let progressoTotal = 0;
    if (proc.etapas.length > 0) {
        proc.etapas.forEach(etapa => {
            progressoTotal += (etapa.progresso || 0) * (etapa.peso || 0.15);
        });
    }
    const progressoPct = Math.round(progressoTotal * 100);
    
    const concluidas = proc.etapas.filter(e => e.status === 'Concluída').length;
    const emExec = proc.etapas.filter(e => e.status === 'Em execução').length;
    const totalEtapas = proc.etapas.length;
    
    let statusGeral = 'Não iniciada';
    let statusClass = 'status-pendente';
    if (concluidas === totalEtapas && totalEtapas > 0) {
        statusGeral = 'Concluída';
        statusClass = 'status-concluida';
    } else if (emExec > 0) {
        statusGeral = 'Em execução';
        statusClass = 'status-em-execucao';
    } else if (concluidas > 0) {
        statusGeral = 'Em andamento';
        statusClass = 'status-em-andamento';
    }
    
    // Calcular duração
    let duracao = '-';
    if (proc.dataInicio && proc.dataTermino) {
        const inicio = new Date(proc.dataInicio);
        const termino = new Date(proc.dataTermino);
        const dias = Math.ceil((termino - inicio) / (1000 * 60 * 60 * 24));
        if (!isNaN(dias)) duracao = dias + ' dias';
    }
    
    // Pegar responsáveis únicos (apenas valores válidos, excluindo status e headers)
    const valoresInvalidos = [
        'Status', 'Em execução', 'Concluída', 'Não iniciada', 'Responsável', '-', '',
        // Variações de status que podem aparecer
        'Concluído', 'Não iniciado', 'Em andamento', 'Paralisado', 'Paralisada',
        'Aguardando', 'Bloqueado', 'Bloqueada', 'Cancelado', 'Cancelada'
    ];
    const responsaveis = [...new Set(
        proc.etapas
            .map(e => e.responsavel)
            .filter(r => r && r.trim() !== '' && !valoresInvalidos.includes(r.trim()))
    )].join(', ') || 'Não definido';
    
    // Título do card: prioriza nome da aba se tiver, senão descrição
    let tituloCard = proc.nome || `Processo ${index + 1}`;
    // Remove "Processo:" se existir
    if (tituloCard.includes('Processo:')) {
        tituloCard = tituloCard.split('Processo:')[1].trim();
    }
    // Se não houver nome descritivo, usa a descrição
    if (tituloCard === `Processo ${index + 1}` && proc.descricao) {
        tituloCard = proc.descricao;
    }
    
    card.innerHTML = `
        <div class="card-header" style="background: linear-gradient(135deg, #1F4E78 0%, #366092 100%)">
            <div class="card-title">
                <span class="processo-id">#${index + 1}</span>
                <h3>${tituloCard}</h3>
            </div>
            <div class="card-actions">
                <button class="btn-expand" onclick="expandProcesso(${index + 1})">
                    <i class="fas fa-chevron-down"></i> Expandir Detalhes
                </button>
            </div>
        </div>
        
        <div class="card-body">
            <div class="card-metrics">
                <div class="metric">
                    <span class="metric-label">Status</span>
                    <span class="status-badge ${statusClass}">${statusGeral}</span>
                </div>
                <div class="metric">
                    <span class="metric-label">% Conclusão</span>
                    <div class="progress-bar">
                        <div class="progress-fill" style="width: ${progressoPct}%">${progressoPct}%</div>
                    </div>
                </div>
                <div class="metric">
                    <span class="metric-label">Etapas</span>
                    <span class="metric-value">${concluidas}/${totalEtapas} <i class="fas fa-tasks"></i></span>
                </div>
                <div class="metric">
                    <span class="metric-label">Responsáveis</span>
                    <span class="metric-value">${responsaveis}</span>
                </div>
            </div>
            
            <div class="card-timeline">
                <div class="timeline-item">
                    <i class="fas fa-calendar-alt"></i>
                    <span><strong>Início:</strong> ${proc.dataInicio || '-'}</span>
                </div>
                <div class="timeline-item">
                    <i class="fas fa-calendar-check"></i>
                    <span><strong>Término:</strong> ${proc.dataTermino || '-'}</span>
                </div>
                <div class="timeline-item">
                    <i class="fas fa-clock"></i>
                    <span><strong>Duração:</strong> ${proc.duracaoTotal || duracao}</span>
                </div>
                <div class="timeline-item">
                    <i class="fas fa-hourglass-half"></i>
                    <span><strong>Dias Restantes:</strong> ${proc.diasRestantes || '-'}</span>
                </div>
            </div>
            
            <div class="card-timeline mt-2">
                <div class="timeline-item">
                    <i class="fas fa-building"></i>
                    <span><strong>Unidade:</strong> ${proc.unidadeDemandante || '-'}</span>
                </div>
                <div class="timeline-item">
                    <i class="fas fa-user-tie"></i>
                    <span><strong>Responsável:</strong> ${proc.responsavelDemanda || responsaveis}</span>
                </div>
                <div class="timeline-item">
                    <i class="fas fa-layer-group"></i>
                    <span><strong>Categoria:</strong> ${proc.categoria || '-'}</span>
                </div>
                ${proc.statusAtual ? `
                <div class="timeline-item">
                    <i class="fas fa-info-circle"></i>
                    <span><strong>Status:</strong> ${proc.statusAtual}</span>
                </div>` : ''}
            </div>
            
            <div class="card-indicators">
                <div class="indicator">
                    <span class="indicator-label">Prioridade</span>
                    <span class="badge badge-${proc.prioridade?.toLowerCase() || 'media'}">${proc.prioridade || 'Média'}</span>
                </div>
                <div class="indicator">
                    <span class="indicator-label">SEI</span>
                    <span class="indicator-value">${proc.sei || '-'}</span>
                </div>
            </div>
        </div>
        
        <div class="card-expanded" id="processo-${index + 1}-details" style="display: none;">
            <div class="expanded-content">
                <h4><i class="fas fa-list-check"></i> Resumo do Processo</h4>
                
                <!-- Etapas com Status Visual -->
                <div class="resumo-section">
                    <strong class="resumo-title">📊 Etapas (${proc.etapas.length})</strong>
                    <div class="etapas-mini">
                        ${proc.etapas.map(etapa => {
                            const statusIcon = etapa.status === 'Concluída' || etapa.status === 'Concluído' ? '✅' : 
                                             etapa.status === 'Em execução' ? '🔄' : 
                                             etapa.status === 'Paralisado' || etapa.status === 'Paralisada' ? '⏸️' : '⚪';
                            return `
                            <div class="etapa-mini etapa-${etapa.status.toLowerCase().replace(' ', '-')}">
                                <div class="etapa-info">
                                    <span class="etapa-icon">${statusIcon}</span>
                                    <span class="etapa-nome">${etapa.nome}</span>
                                    ${etapa.responsavel ? `<span class="etapa-resp">👤 ${etapa.responsavel}</span>` : ''}
                                </div>
                                <div class="etapa-metrics">
                                    <span class="etapa-progress">${Math.round((etapa.progresso || 0) * 100)}%</span>
                                    ${etapa.dataTermino ? `<span class="etapa-prazo">📅 ${etapa.dataTermino}</span>` : ''}
                                </div>
                            </div>
                        `;
                        }).join('')}
                    </div>
                </div>
                
                <!-- Tarefas Agrupadas por Etapa -->
                ${proc.tarefas && proc.tarefas.length > 0 ? `
                <div class="resumo-section">
                    <strong class="resumo-title">✅ Tarefas (${proc.tarefas.length})</strong>
                    <div class="tarefas-mini">
                        ${(() => {
                            // Agrupar tarefas por etapa
                            const tarefasPorEtapa = {};
                            proc.tarefas.forEach(tarefa => {
                                const etapa = tarefa.etapa || 'Sem etapa';
                                if (!tarefasPorEtapa[etapa]) {
                                    tarefasPorEtapa[etapa] = [];
                                }
                                tarefasPorEtapa[etapa].push(tarefa);
                            });
                            
                            // Renderizar tarefas agrupadas
                            return Object.entries(tarefasPorEtapa).map(([etapa, tarefas]) => `
                                <div class="tarefa-grupo">
                                    <div class="tarefa-etapa-nome">📁 ${etapa}</div>
                                    ${tarefas.map(tarefa => {
                                        const statusIcon = tarefa.status === 'Concluída' || tarefa.status === 'Concluído' ? '✅' : 
                                                         tarefa.status === 'Em execução' ? '🔄' : '⚪';
                                        return `
                                        <div class="tarefa-mini">
                                            <span class="tarefa-icon">${statusIcon}</span>
                                            <span class="tarefa-nome">${tarefa.nome}</span>
                                            <span class="tarefa-progress">${Math.round((tarefa.progresso || 0) * 100)}%</span>
                                        </div>
                                        `;
                                    }).join('')}
                                </div>
                            `).join('');
                        })()}
                    </div>
                </div>
                ` : ''}
            </div>
        </div>
    `;
    
    return card;
}

// ==================== CRIAR ABAS DINÂMICAS PARA PROCESSOS ====================
function criarAbasProcessos() {
    const tabsNav = document.getElementById('tabs-navigation');
    const dashboardBtn = tabsNav.querySelector('[data-tab="dashboard"]');
    const container = document.getElementById('processos-tabs-container');
    
    if (!tabsNav || !container) return;
    
    // Remover abas antigas de processos
    tabsNav.querySelectorAll('.tab-btn-processo').forEach(btn => btn.remove());
    container.innerHTML = '';
    
    // Criar aba e conteúdo para cada processo
    todosProcessos.forEach((proc, index) => {
        const processoId = `processo-${index + 1}`;
        const processoNum = index + 1;
        
        // Título para a aba de navegação
        let tituloAba = proc.nome || `Processo ${processoNum}`;
        
        // Se o nome da aba contém "Processo:", pega só o texto depois
        if (tituloAba.includes('Processo:')) {
            tituloAba = tituloAba.split('Processo:')[1].trim();
        }
        
        // Se o título for muito longo (>30 caracteres), abrevia
        if (tituloAba.length > 30) {
            tituloAba = tituloAba.substring(0, 27) + '...';
        }
        
        // Criar botão de aba
        const tabBtn = document.createElement('button');
        tabBtn.className = 'tab-btn tab-btn-processo';
        tabBtn.setAttribute('data-tab', processoId);
        tabBtn.innerHTML = `📂 ${tituloAba}`;
        tabBtn.onclick = () => switchTab(processoId);
        
        // Inserir antes do Dashboard
        dashboardBtn.before(tabBtn);
        
        // Criar conteúdo da aba
        const tabContent = criarConteudoAbaProcesso(proc, index);
        container.appendChild(tabContent);
    });
}

function criarConteudoAbaProcesso(proc, index) {
    const processoId = `processo-${index + 1}`;
    const processoNum = index + 1;
    
    // Calcular métricas
    let progressoTotal = 0;
    if (proc.etapas.length > 0) {
        proc.etapas.forEach(etapa => {
            progressoTotal += (etapa.progresso || 0) * (etapa.peso || 0.15);
        });
    }
    const progressoPct = Math.round(progressoTotal * 100);
    
    const concluidas = proc.etapas.filter(e => e.status === 'Concluída').length;
    const emExec = proc.etapas.filter(e => e.status === 'Em execução').length;
    const totalEtapas = proc.etapas.length;
    
    let statusGeral = 'Não iniciada';
    if (concluidas === totalEtapas && totalEtapas > 0) {
        statusGeral = 'Concluída';
    } else if (emExec > 0) {
        statusGeral = 'Em execução';
    } else if (concluidas > 0) {
        statusGeral = 'Em andamento';
    }
    
    // Calcular duração
    let duracao = '-';
    if (proc.dataInicio && proc.dataTermino) {
        const inicio = new Date(proc.dataInicio);
        const termino = new Date(proc.dataTermino);
        const dias = Math.ceil((termino - inicio) / (1000 * 60 * 60 * 24));
        if (!isNaN(dias)) duracao = dias + ' dias';
    }
    
    const section = document.createElement('section');
    section.id = processoId;
    section.className = 'tab-content';
    
    // Título da aba: usa descrição se disponível, senão usa nome da aba
    const tituloProcesso = proc.descricao || proc.nome || `Processo ${processoNum}`;
    
    section.innerHTML = `
        <div class="processo-header">
            <h2 class="processo-title">📂 ${tituloProcesso}</h2>
        </div>

        <div class="projeto-info">
            <h3 class="subsection-title">📋 INFORMAÇÕES DO PROJETO</h3>
            <div class="info-grid">
                <div class="info-item">
                    <span class="info-label">SEI:</span>
                    <span class="info-value">${proc.sei || '-'}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Unidade:</span>
                    <span class="info-value">${proc.unidadeDemandante || '-'}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Categoria:</span>
                    <span class="info-value">${proc.categoria || '-'}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Prioridade:</span>
                    <span class="badge badge-${proc.prioridade?.toLowerCase() || 'media'}">${proc.prioridade || 'Média'}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Responsável:</span>
                    <span class="info-value">${proc.responsavelDemanda || '-'}</span>
                </div>
            </div>
            
            <div class="info-grid mt-2">
                <div class="info-item">
                    <span class="info-label">Data Início:</span>
                    <span class="info-value">${proc.dataInicio || '-'}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Data Término:</span>
                    <span class="info-value">${proc.dataTermino || '-'}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Duração Total:</span>
                    <span class="info-value">${proc.duracaoTotal || duracao}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Dias Restantes:</span>
                    <span class="info-value ${parseInt(proc.diasRestantes) < 0 ? 'text-danger' : ''}">${proc.diasRestantes || '-'}</span>
                </div>
            </div>

            <div class="info-grid mt-2">
                <div class="info-item highlight">
                    <span class="info-label">Progresso Geral:</span>
                    <span class="info-value-big">${proc.progressoGeral || progressoPct + '%'}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Status Atual:</span>
                    <span class="status-badge status-${(proc.statusAtual || statusGeral).toLowerCase().replace(' ', '-')}">${proc.statusAtual || statusGeral}</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Etapas:</span>
                    <span class="info-value">${concluidas}/${totalEtapas} concluídas</span>
                </div>
            </div>

            <div class="descricao-box">
                <strong>Descrição:</strong>
                <p>${proc.descricao || 'Sem descrição disponível'}</p>
            </div>
        </div>

        <!-- SUB-NAVEGAÇÃO: ETAPAS E TAREFAS -->
        <div class="sub-tabs-navigation">
            <button class="sub-tab-btn active" onclick="switchSubTab('${processoId}', 'etapas')">
                🔄 ETAPAS DO PROCESSO
            </button>
            <button class="sub-tab-btn" onclick="switchSubTab('${processoId}', 'tarefas')">
                📝 TAREFAS DETALHADAS
            </button>
        </div>

        <!-- CONTEÚDO: ETAPAS -->
        <div id="${processoId}-etapas" class="sub-tab-content active">
            <div class="etapas-section">
                <h3 class="subsection-title">🔄 ETAPAS DO PROCESSO - DETALHAMENTO COMPLETO</h3>
                <p class="hint">💡 Total de ${totalEtapas} etapa(s) - ${concluidas} concluída(s), ${emExec} em execução</p>
                
                <div class="table-responsive">
                    <table class="etapas-table">
                        <thead>
                            <tr>
                                <th style="width: 20%">Etapa</th>
                                <th style="width: 10%">Status</th>
                                <th style="width: 12%">Responsável</th>
                                <th style="width: 8%">Início</th>
                                <th style="width: 8%">Término</th>
                                <th style="width: 10%">Situação</th>
                                <th style="width: 12%">Progresso</th>
                                <th style="width: 8%">Horas</th>
                                <th style="width: 6%">Peso</th>
                                <th style="width: 6%">Ações</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${proc.etapas.map((etapa, etapaIdx) => `
                                <tr>
                                    <td><strong>${etapa.nome}</strong></td>
                                    <td><span class="status-badge status-${etapa.status.toLowerCase().replace(' ', '-')}">${etapa.status}</span></td>
                                    <td>${etapa.responsavel || '-'}</td>
                                    <td>${etapa.dataInicio || '-'}</td>
                                    <td>${etapa.dataTermino || '-'}</td>
                                    <td><span class="badge badge-situacao ${etapa.situacao ? '' : 'hidden'}">${etapa.situacao || '-'}</span></td>
                                    <td>
                                        <div class="progress-bar">
                                            <div class="progress-fill" style="width: ${(etapa.progresso || 0) * 100}%">
                                                ${Math.round((etapa.progresso || 0) * 100)}%
                                            </div>
                                        </div>
                                    </td>
                                    <td>${etapa.horasEstimadas || 0}h / ${etapa.horasReais || 0}h</td>
                                    <td>${((etapa.peso || 0) * 100).toFixed(0)}%</td>
                                    <td>
                                        ${etapa.tarefasTexto || etapa.observacoes || etapa.produtos ? 
                                            `<button class="btn-mini" onclick="toggleEtapaDetails('${processoId}-etapa-${etapaIdx}')"><i class="fas fa-chevron-down"></i></button>` 
                                            : '<span style="color: #999">-</span>'}
                                    </td>
                                </tr>
                                ${etapa.tarefasTexto || etapa.observacoes || etapa.produtos ? `
                                <tr id="${processoId}-etapa-${etapaIdx}" class="etapa-details" style="display: none;">
                                    <td colspan="10">
                                        <div class="etapa-details-content">
                                            ${etapa.produtos ? `
                                            <div class="detail-section">
                                                <strong><i class="fas fa-box"></i> Produtos/Entregas:</strong>
                                                <p>${etapa.produtos}</p>
                                            </div>` : ''}
                                            ${etapa.tarefasTexto ? `
                                            <div class="detail-section">
                                                <strong><i class="fas fa-tasks"></i> Tarefas:</strong>
                                                <p style="white-space: pre-line;">${etapa.tarefasTexto}</p>
                                            </div>` : ''}
                                            ${etapa.observacoes ? `
                                            <div class="detail-section">
                                                <strong><i class="fas fa-comment-alt"></i> Observações:</strong>
                                                <p style="white-space: pre-line;">${etapa.observacoes}</p>
                                            </div>` : ''}
                                            ${etapa.dependencias && etapa.dependencias !== '-' ? `
                                            <div class="detail-section">
                                                <strong><i class="fas fa-link"></i> Dependências:</strong>
                                                <p>${etapa.dependencias}</p>
                                            </div>` : ''}
                                        </div>
                                    </td>
                                </tr>` : ''}
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- CONTEÚDO: TAREFAS -->
        <div id="${processoId}-tarefas" class="sub-tab-content">
            <div class="tarefas-section">
                <h3 class="subsection-title">📝 TAREFAS DETALHADAS POR ETAPA</h3>
            
            <div class="table-responsive">
                <table class="tarefas-table">
                    <thead>
                        <tr>
                            <th>Etapa</th>
                            <th>Tarefa</th>
                            <th>Status</th>
                            <th>Responsável</th>
                            <th>Prioridade</th>
                            <th>Prazo</th>
                            <th>% Conclusão</th>
                            <th>Horas</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${proc.tarefas.map(tarefa => `
                            <tr>
                                <td>${tarefa.etapa || '-'}</td>
                                <td><strong>${tarefa.nome}</strong></td>
                                <td><span class="status-badge status-${tarefa.status.toLowerCase().replace(' ', '-')}">${tarefa.status}</span></td>
                                <td>${tarefa.responsavel || '-'}</td>
                                <td><span class="badge badge-${tarefa.prioridade?.toLowerCase() || 'media'}">${tarefa.prioridade || 'Média'}</span></td>
                                <td>${tarefa.prazo || '-'}</td>
                                <td>${Math.round((tarefa.progresso || 0) * 100)}%</td>
                                <td>${tarefa.horas || 0}h</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>
        </div>
    `;
    
    return section;
}

// ==================== NAVEGAÇÃO DE SUB-ABAS ====================
function switchSubTab(processoId, subtab) {
    // Remover active de todos os botões e conteúdos do processo específico
    const processoSection = document.getElementById(processoId);
    if (!processoSection) return;
    
    processoSection.querySelectorAll('.sub-tab-btn').forEach(btn => btn.classList.remove('active'));
    processoSection.querySelectorAll('.sub-tab-content').forEach(content => content.classList.remove('active'));
    
    // Adicionar active aos selecionados
    const targetBtn = processoSection.querySelector(`.sub-tab-btn[onclick*="${subtab}"]`);
    const targetContent = document.getElementById(`${processoId}-${subtab}`);
    
    if (targetBtn) targetBtn.classList.add('active');
    if (targetContent) targetContent.classList.add('active');
}

// ==================== TOGGLE DE DETALHES DAS ETAPAS ====================
function toggleEtapaDetails(etapaId) {
    const detailsRow = document.getElementById(etapaId);
    const btn = event.target.closest('.btn-mini');
    
    if (detailsRow) {
        const isHidden = detailsRow.style.display === 'none';
        detailsRow.style.display = isHidden ? 'table-row' : 'none';
        
        if (btn) {
            const icon = btn.querySelector('i');
            if (icon) {
                icon.className = isHidden ? 'fas fa-chevron-up' : 'fas fa-chevron-down';
            }
        }
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
    
    // Não iniciar auto-refresh se o intervalo for 0 (desativado)
    if (GOOGLE_SHEETS_CONFIG.autoRefreshInterval === 0) {
        console.log('⏸️ Auto-refresh desativado. Use o botão "Atualizar" para refresh manual.');
        return;
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
            switchTab(tabId);
        });
    });
}

function switchTab(tabId) {
    // Remover classe active de todos os botões e conteúdos
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    
    // Adicionar classe active ao botão e conteúdo selecionados
    const targetBtn = document.querySelector(`[data-tab="${tabId}"]`);
    const targetContent = document.getElementById(tabId);
    
    if (targetBtn) targetBtn.classList.add('active');
    if (targetContent) targetContent.classList.add('active');
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

function calcularKPIsGlobais() {
    if (todosProcessos.length === 0) {
        return calcularKPIs(); // Fallback para dados locais
    }
    
    let totalProcessos = todosProcessos.length;
    let processosAtivos = 0;
    let processosConcluidos = 0;
    let processosPlanejados = 0;
    let processosAtrasados = 0;
    let somaProgresso = 0;
    let somaDias = 0;
    
    todosProcessos.forEach(proc => {
        // Calcular progresso do processo
        let progressoProcesso = 0;
        if (proc.etapas.length > 0) {
            proc.etapas.forEach(etapa => {
                progressoProcesso += (etapa.progresso || 0) * (etapa.peso || 0.15);
            });
        }
        
        somaProgresso += progressoProcesso;
        
        // Contar etapas por status
        const emExec = proc.etapas.filter(e => e.status === 'Em execução').length;
        const concluidas = proc.etapas.filter(e => e.status === 'Concluída').length;
        const naoIniciadas = proc.etapas.filter(e => e.status === 'Não iniciada').length;
        
        if (emExec > 0) processosAtivos++;
        if (concluidas === proc.etapas.length && proc.etapas.length > 0) processosConcluidos++;
        if (naoIniciadas === proc.etapas.length && proc.etapas.length > 0) processosPlanejados++;
        
        // Calcular dias
        if (proc.dataInicio && proc.dataTermino) {
            const inicio = new Date(proc.dataInicio);
            const termino = new Date(proc.dataTermino);
            const dias = Math.ceil((termino - inicio) / (1000 * 60 * 60 * 24));
            if (!isNaN(dias)) somaDias += dias;
        }
    });
    
    const progressoMedio = totalProcessos > 0 ? Math.round((somaProgresso / totalProcessos) * 100) : 0;
    const prazoMedio = totalProcessos > 0 ? Math.round(somaDias / totalProcessos) : 0;
    
    // Atualizar KPIs
    const kpiTotal = document.getElementById('kpi-total');
    const kpiAtivos = document.getElementById('kpi-ativos');
    const kpiConcluidos = document.getElementById('kpi-concluidos');
    const kpiPlanejados = document.getElementById('kpi-planejados');
    const kpiAtrasados = document.getElementById('kpi-atrasados');
    const kpiProgresso = document.getElementById('kpi-progresso');
    const kpiPrazo = document.getElementById('kpi-prazo');
    
    if (kpiTotal) kpiTotal.textContent = totalProcessos;
    if (kpiAtivos) kpiAtivos.textContent = processosAtivos;
    if (kpiConcluidos) kpiConcluidos.textContent = processosConcluidos;
    if (kpiPlanejados) kpiPlanejados.textContent = processosPlanejados;
    if (kpiAtrasados) kpiAtrasados.textContent = processosAtrasados;
    if (kpiProgresso) kpiProgresso.textContent = progressoMedio + '%';
    if (kpiPrazo) kpiPrazo.textContent = prazoMedio + 'd';
    
    // Atualizar alertas
    const alertsDiv = document.getElementById('alerts-container');
    if (alertsDiv) {
        alertsDiv.innerHTML = '';
        
        if (processosAtivos > 0) {
            alertsDiv.innerHTML += `<div class="alert alert-warning">🟡 ATENÇÃO: ${processosAtivos} processo(s) em execução</div>`;
        }
        
        if (processosConcluidos > 0) {
            alertsDiv.innerHTML += `<div class="alert alert-success">🟢 SUCESSO: ${processosConcluidos} processo(s) concluído(s)!</div>`;
        }
        
        if (progressoMedio >= 70) {
            alertsDiv.innerHTML += '<div class="alert alert-success">🟢 Progresso geral excelente!</div>';
        } else if (progressoMedio < 30) {
            alertsDiv.innerHTML += '<div class="alert alert-warning">🟡 Progresso geral baixo - atenção necessária</div>';
        }
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
