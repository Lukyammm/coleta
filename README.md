# Clinify - Coleta de Dados Clínicos

Um Web App moderno e de alta performance construído nativamente em **Google Apps Script** para coleta e gerenciamento de dados rápidos (nome, prontuário, justificativa). O projeto utiliza uma arquitetura SPA (Single Page Application) em Vanilla JS, inspirada nas melhores práticas de design de sistemas.

## 🌟 Funcionalidades

- **Design System Profissional:** UI baseada em tokens/variáveis CSS rigorosas, garantindo consistência visual de alto nível.
- **Microinterações Sonoras (UX/Psicologia):** Uso da *Web Audio API* para emitir frequências sutis ao focar em inputs, enviar dados e receber sucesso, ativando recompensas de dopamina (sem uso pesado de base64).
- **Compatibilidade Extrema (Mobile-first):** Suporte nativo ao `safe-area-inset` do iOS/Safari, sem atrasos de clique e sem zoom forçado em inputs (`user-scalable=no`).
- **Arquitetura RPC (Remote Procedure Call):** Comunicação elegante usando Promises (`async/await`) em cima do `google.script.run` nativo, abolindo a necessidade de aninhar callbacks (*callback-hell*).
- **Dashboard Administrativo Seguro:** Acesso bloqueado por senha (validada no back-end) com visualização em Grid responsivo dos dados coletados em tempo real (calcula os registros de "Hoje" e formatação dinâmica de timestamps).

## 🛠️ Tecnologias

- **Front-end:** HTML5, Vanilla JavaScript (ES6+), CSS3 Puro.
- **Back-end & Banco de Dados:** Google Apps Script (Javascript/V8 Engine) rodando sob Google Sheets.

## 🚀 Como fazer o Deploy

1. Crie uma nova Planilha no Google Sheets (ela servirá como banco de dados).
2. No menu da planilha, vá em **Extensões > Apps Script**.
3. No painel que abrir, copie o conteúdo do arquivo `Code.gs` deste repositório e cole sobrepondo o código original.
4. Clique no ícone de **+** (Adicionar um arquivo) > **HTML** e crie um arquivo com o nome **`index`**.
5. Copie o conteúdo de `index.html` deste repositório e cole no arquivo criado.
6. Salve tudo clicando no ícone do disquete.
7. Clique em **Implantar (Deploy) > Nova implantação**.
8. Na engrenagem, escolha **App da Web**.
   - Executar como: *Eu*
   - Quem tem acesso: *Qualquer pessoa* (ou restrito a contas específicas).
9. Conceda as permissões de conta (o Google exibirá um alerta, vá em *Avançado > Acessar o projeto*).
10. Copie a URL do App da Web e abra no celular.

> **Nota:** A senha inicial padrão para acessar a área administrativa no App da Web é `adm`.