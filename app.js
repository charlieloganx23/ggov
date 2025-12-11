// ==================== DADOS DO SISTEMA ====================
const processoData = {
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

// ==================== INICIALIZAÇÃO ====================
document.addEventListener('DOMContentLoaded', function() {
    updateTimestamp();
    setInterval(updateTimestamp, 1000);
    
    initTabs();
    renderEtapas();
    renderTarefas();
    initCharts();
    calcularKPIs();
});

// ==================== TIMESTAMP ====================
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
    document.getElementById('timestamp').textContent = formatted;
}

// ==================== TABS ====================
function initTabs() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    
    tabBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            const tabId = this.dataset.tab;
            
            // Remove active de todos
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            
            // Adiciona active ao clicado
            this.classList.add('active');
            document.getElementById(tabId).classList.add('active');
        });
    });
}

// ==================== EXPANDIR PROCESSO ====================
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

// ==================== RENDERIZAR ETAPAS ====================
function renderEtapas() {
    const tbody = document.getElementById('etapas-tbody');
    
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

// ==================== RENDERIZAR TAREFAS ====================
function renderTarefas() {
    const tbody = document.getElementById('tarefas-tbody');
    
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

// ==================== FILTRAR TAREFAS ====================
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

// ==================== CHARTS ====================
function initCharts() {
    // Chart 1: Status das Etapas
    const statusCtx = document.getElementById('statusChart').getContext('2d');
    
    const statusCount = {
        'Em execução': 0,
        'Concluída': 0,
        'Não iniciada': 0
    };
    
    processoData.etapas.forEach(etapa => {
        statusCount[etapa.status]++;
    });
    
    new Chart(statusCtx, {
        type: 'doughnut',
        data: {
            labels: Object.keys(statusCount),
            datasets: [{
                data: Object.values(statusCount),
                backgroundColor: [
                    '#FFA500',
                    '#70AD47',
                    '#5B9BD5'
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    position: 'bottom'
                }
            }
        }
    });
    
    // Chart 2: Progresso por Etapa
    const progressCtx = document.getElementById('progressChart').getContext('2d');
    
    new Chart(progressCtx, {
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
            scales: {
                y: {
                    beginAtZero: true,
                    max: 100
                }
            },
            plugins: {
                legend: {
                    display: false
                }
            }
        }
    });
}

// ==================== CALCULAR KPIs ====================
function calcularKPIs() {
    // Calcular % médio ponderado
    let somaProgresso = 0;
    processoData.etapas.forEach(etapa => {
        somaProgresso += etapa.progresso * etapa.peso;
    });
    
    const percentualGeral = Math.round(somaProgresso * 100);
    
    // Atualizar KPIs
    const emExecucao = processoData.etapas.filter(e => e.status === 'Em execução').length;
    const concluidas = processoData.etapas.filter(e => e.status === 'Concluída').length;
    const naoIniciadas = processoData.etapas.filter(e => e.status === 'Não iniciada').length;
    
    document.getElementById('kpi-ativos').textContent = emExecucao;
    document.getElementById('kpi-concluidos').textContent = concluidas;
    document.getElementById('kpi-planejados').textContent = naoIniciadas;
    document.getElementById('kpi-saude').textContent = percentualGeral + '%';
    
    // Calcular horas
    const horasTotal = processoData.etapas.reduce((sum, e) => sum + e.horasEstimadas, 0);
    const horasReais = processoData.etapas.reduce((sum, e) => sum + e.horasReais, 0);
    
    // Atualizar alertas
    const alertsDiv = document.getElementById('alerts-container');
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

// ==================== HELPERS ====================
function formatDate(dateString) {
    const date = new Date(dateString);
    return date.toLocaleDateString('pt-BR');
}

function verDetalhesEtapa(index) {
    alert('Funcionalidade em desenvolvimento!\nEtapa: ' + processoData.etapas[index].nome);
}

// ==================== ANIMAÇÕES ====================
window.addEventListener('scroll', function() {
    const cards = document.querySelectorAll('.processo-card, .chart-card, .stats-card');
    
    cards.forEach(card => {
        const cardTop = card.getBoundingClientRect().top;
        const windowHeight = window.innerHeight;
        
        if (cardTop < windowHeight - 100) {
            card.style.opacity = '1';
            card.style.transform = 'translateY(0)';
        }
    });
});

console.log('🚀 Sistema GGOV Revolucionário carregado com sucesso!');
console.log('📊 Dados do processo:', processoData);
