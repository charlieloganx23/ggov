# 📚 Guia: Como Adicionar Múltiplos Processos

## 🎯 Visão Geral

O sistema agora suporta **múltiplos processos** simultaneamente! Cada processo é uma aba na planilha do Google Sheets.

---

## 📋 Como Adicionar um Novo Processo

### **Passo 1: Duplicar a Aba no Google Sheets**

1. Abra sua planilha no Google Sheets
2. Clique com botão direito na aba **"Processo 1"**
3. Selecione **"Duplicar"**
4. Renomeie a nova aba para **"Processo 2"** (ou qualquer nome)

### **Passo 2: Preencher os Dados do Novo Processo**

Na nova aba, preencha:

- **Linha 3**: Informações do projeto
  - SEI, Prioridade, Categoria, Data Início, Data Término, Descrição

- **Linha 6+**: Etapas
  - Nome, Status, Responsável, Datas, Produtos, etc.

- **Linha 16+**: Tarefas
  - Etapa, Tarefa, Status, Responsável, Prioridade, etc.

### **Passo 3: Configurar no Sistema**

Edite o arquivo `app-google-sheets.js`, na linha **~15**:

```javascript
processos: [
    'Processo 1',
    'Processo 2',  // ← Adicione aqui!
    'Processo 3',  // ← E aqui!
    // ... adicione quantos quiser
],
```

### **Passo 4: Salvar e Atualizar**

1. Salve o arquivo `app-google-sheets.js`
2. Faça commit no Git:
   ```powershell
   git add .
   git commit -m "feat: Adiciona Processo 2"
   git push
   ```

3. Aguarde deploy automático no Netlify (1-2 minutos)

---

## ✨ O que Acontece Automaticamente

O sistema irá:

✅ **Detectar** todos os processos configurados  
✅ **Carregar** dados de cada aba do Google Sheets  
✅ **Criar** cards dinâmicos para cada processo  
✅ **Calcular** KPIs globais (somando todos os processos)  
✅ **Atualizar** alertas e notificações  

---

## 📊 KPIs Globais

Os indicadores mostram dados **consolidados** de todos os processos:

- **Total Processos**: Quantidade de processos monitorados
- **Em Execução**: Processos com pelo menos 1 etapa em execução
- **Concluídos**: Processos com todas as etapas concluídas
- **Planejados**: Processos que ainda não iniciaram
- **Progresso Geral**: Média do progresso de todos os processos
- **Prazo Médio**: Média de duração em dias

---

## 🔄 Sincronização Automática

- Dados atualizados **a cada 30 segundos**
- Todos os processos sincronizam simultaneamente
- Alterações na planilha aparecem automaticamente

---

## 💡 Dicas

1. **Mantenha o mesmo padrão** de estrutura em todas as abas
2. **Use nomes claros** para as abas (Processo 1, Processo 2, etc.)
3. **Preencha todas as colunas obrigatórias** (SEI, Descrição, etc.)
4. **Evite caracteres especiais** nos nomes das abas

---

## ⚙️ Estrutura Técnica

```
Google Sheets (1 planilha)
├── Processo 1 (aba)
│   ├── Linha 3: Info do projeto
│   ├── Linhas 6+: Etapas
│   └── Linhas 16+: Tarefas
│
├── Processo 2 (aba)
│   ├── Linha 3: Info do projeto
│   ├── Linhas 6+: Etapas
│   └── Linhas 16+: Tarefas
│
└── ... (quantas abas quiser)
```

---

## 🚀 Exemplo Prático

**Cenário**: Você gerencia 3 processos GGOV

1. Crie 3 abas no Google Sheets:
   - "Processo 1" - Mapeamento de processos
   - "Processo 2" - Capacitação de servidores
   - "Processo 3" - Implementação de sistema

2. Configure em `app-google-sheets.js`:
   ```javascript
   processos: [
       'Processo 1',
       'Processo 2',
       'Processo 3'
   ],
   ```

3. O sistema mostrará:
   - **3 cards** no Command Center
   - **KPIs consolidados** dos 3 processos
   - **Alertas** relevantes de todos

---

## 📞 Precisa de Ajuda?

Se tiver dúvidas ou problemas:
1. Verifique o console do navegador (F12)
2. Confirme que os nomes das abas estão corretos
3. Valide se a estrutura das linhas está mantida

---

**Última atualização**: 11/12/2025  
**Versão do sistema**: 2.0 (Multi-processos)
