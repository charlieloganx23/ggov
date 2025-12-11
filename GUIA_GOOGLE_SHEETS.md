# 🚀 Sistema GGOV - Integração com Google Sheets

## 📋 Guia de Configuração Completo

### ✨ O que foi implementado:

✅ **Sincronização bidirecional** com Google Sheets  
✅ **Atualização automática** a cada 30 segundos  
✅ **Botões de controle**: Atualizar manual, Ativar/Pausar auto-refresh  
✅ **Indicador de status** de conexão em tempo real  
✅ **Modal de configuração** amigável  
✅ **Notificações visuais** para todas as ações  
✅ **Fallback para dados locais** se desconectado  
✅ **Persistência** de configuração no localStorage  

---

## 🔧 Passo a Passo para Configurar

### **1️⃣ Criar Projeto no Google Cloud**

1. Acesse: https://console.cloud.google.com/
2. Clique em **"Selecionar projeto"** → **"Novo Projeto"**
3. Nome do projeto: `Sistema GGOV`
4. Clique em **"Criar"**

### **2️⃣ Ativar Google Sheets API**

1. No menu lateral, vá em **"APIs e Serviços"** → **"Biblioteca"**
2. Pesquise por: `Google Sheets API`
3. Clique no resultado e depois em **"Ativar"**

### **3️⃣ Criar API Key**

1. No menu lateral, vá em **"APIs e Serviços"** → **"Credenciais"**
2. Clique em **"+ Criar Credenciais"** → **"Chave de API"**
3. Uma API Key será gerada (exemplo: `AIzaSyXXXXXXXXXXXXXXXXXXXXXX`)
4. **COPIE** esta chave (você vai precisar dela!)

### **4️⃣ Configurar Restrições da API Key (Importante!)**

1. Na tela de credenciais, clique no nome da API Key criada
2. Em **"Restrições da API"**, selecione **"Restringir chave"**
3. Marque apenas: **Google Sheets API**
4. Clique em **"Salvar"**

### **5️⃣ Preparar sua Planilha Google Sheets**

1. Abra o arquivo Excel gerado: `Sistema_GGOV_Revolucionario.xlsx`
2. Faça upload para Google Drive
3. Abra com Google Sheets
4. **IMPORTANTE**: Torne a planilha **pública** ou compartilhada:
   - Clique em **"Compartilhar"** (canto superior direito)
   - Em **"Obter link"**, selecione: **"Qualquer pessoa com o link"** → **"Leitor"**
   - Clique em **"Concluído"**

### **6️⃣ Copiar ID da Planilha**

Na URL da planilha, copie o ID:

```
https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
                                       ↑_________________________________________↑
                                                  Este é o SPREADSHEET ID
```

### **7️⃣ Configurar o Sistema Web**

1. Abra o arquivo `index.html` no navegador
2. Clique no botão **⚙️** (Configurar)
3. Cole sua **API Key** no primeiro campo
4. Cole o **Spreadsheet ID** no segundo campo
5. Clique em **"Salvar e Conectar"**

---

## 🎯 Funcionalidades Implementadas

### **Botões de Controle:**

🔄 **Atualizar** - Sincroniza dados manualmente com a planilha  
⏰ **Auto (30s)** - Liga/Desliga atualização automática a cada 30 segundos  
🟢 **Status** - Mostra se está conectado ao Google Sheets  
⚙️ **Configurar** - Abre modal de configuração  

### **Como Usar:**

1. **Edite dados na planilha** Google Sheets (altere status, %, horas, etc.)
2. **Aguarde 30 segundos** (auto-refresh) ou clique em **"Atualizar"**
3. **Veja as mudanças** refletidas automaticamente no sistema web!

---

## 📊 Estrutura da Planilha Esperada

O sistema espera encontrar na aba **"Processo 1"**:

### **Etapas (Linhas 12-17):**
- Coluna A: Nome da etapa
- Coluna B: Status (Em execução, Concluída, Não iniciada)
- Coluna C: Responsável
- Coluna D: Data Início
- Coluna E: Data Término
- Coluna F: Produtos/Entregas
- Coluna G: Dependências
- Coluna H: % Progresso (0.0 a 1.0)
- Coluna I: Horas Estimadas
- Coluna J: Horas Reais
- Coluna K: Peso (0.0 a 1.0)

### **Tarefas (Linhas 21-30):**
- Coluna A: Etapa
- Coluna B: Nome da Tarefa
- Coluna C: Status
- Coluna D: Responsável
- Coluna E: Prioridade
- Coluna F: Prazo
- Coluna G: % Conclusão (0.0 a 1.0)
- Coluna H: Horas

---

## 🔐 Segurança

⚠️ **Importante**: 
- A API Key é armazenada no navegador (localStorage)
- A planilha precisa ter permissão de leitura pública
- **NÃO compartilhe** sua API Key publicamente
- Use restrições de API no Google Cloud Console

---

## 🆘 Solução de Problemas

### ❌ "Erro ao conectar com Google Sheets"
- Verifique se a API Key está correta
- Confirme se o Spreadsheet ID está correto
- Certifique-se que a planilha está compartilhada (pública ou com link)
- Verifique se a Google Sheets API está ativada

### ❌ "Nenhum dado encontrado na planilha"
- Confirme que a aba se chama exatamente **"Processo 1"**
- Verifique se os dados estão nas células corretas (A12:L17 para etapas)
- Certifique-se que os headers estão na linha 11

### ⚠️ "Dados não atualizam automaticamente"
- Clique no botão **"Auto (30s)"** para ativar
- Verifique o console do navegador (F12) para erros
- Teste clicando em **"Atualizar"** manualmente

---

## 💡 Dicas

✅ Mantenha a estrutura da planilha Excel original  
✅ Use os status exatos: "Em execução", "Concluída", "Não iniciada"  
✅ Valores de % devem ser decimais (0.7 = 70%)  
✅ O auto-refresh consome menos recursos se pausado quando não estiver usando  

---

## 🚀 Próximos Passos (Possíveis Melhorias)

- [ ] Edição inline no sistema (escrever de volta na planilha)
- [ ] Suporte a múltiplos processos
- [ ] Histórico de alterações
- [ ] Notificações push quando planilha é atualizada
- [ ] Modo offline com sincronização posterior
- [ ] Dashboard de auditoria de mudanças

---

## 📞 Suporte

Se tiver problemas, verifique o **Console do Navegador** (F12 → Console) para ver mensagens de erro detalhadas.

---

**Desenvolvido com ❤️ para o Gabinete de Governança (GGOV)**
