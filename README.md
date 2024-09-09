# RPA Projects Repository #
Este repositório contém diversos projetos de RPA (Automação de Processos Robóticos) desenvolvidos para automatizar tarefas repetitivas e otimizar processos cotidianos. Cada projeto foi criado com o objetivo de facilitar o dia a dia, permitindo que atividades manuais e demoradas sejam realizadas de forma eficiente e precisa.

## Projetos ##

### Envio de E-mails em Massa ###
**Descrição:**
- Este script automatiza o envio de e-mails em massa para destinatários cujos endereços estão listados em uma planilha Excel. Ele utiliza o navegador Microsoft Edge e o serviço de e-mail Outlook para realizar o processo de forma totalmente automatizada.

**Funcionalidades:**
- Leitura de endereços de e-mail de um arquivo Excel.
- Automação do Outlook para criar e enviar e-mails.
- Tratamento de erros e garantia de execução robusta.

**Como usar:**
- Configurar os pré-requisitos (Microsoft Edge, WebDriver, etc.).
- Inserir os endereços de e-mail em uma planilha Excel.
- Executar o script para enviar os e-mails.

**Arquivo:**
- enviar_email_em_massa_excel.py

### Automação de Cadastro de Servidores ###
**Descrição:**
Este script automatiza o processo de atualização e cadastro de servidores em um sistema através do navegador Microsoft Edge. Ele utiliza o Selenium WebDriver para realizar a automação, onde acessa um painel de administração, atualiza informações de servidores listados em uma planilha Excel e marca checkboxes para realizar os cadastros de novos servidores.

**Funcionalidades:**
- Leitura de cadastros (endereços de servidores) a partir de um arquivo Excel.
- Acesso automatizado a um painel web usando o Microsoft Edge.
- Atualização e inserção de novos cadastros de servidores no sistema.
- Tratamento de erros para garantir uma execução robusta, mesmo em caso de falhas de conexão ou tempo limite.

**Como usar:**
1. **Pré-requisitos**:
   - Instalar o Microsoft Edge e o Edge WebDriver.
   - Instalar as bibliotecas Python necessárias: `pandas`, `selenium`, e `openpyxl`.
   
   ```bash
   pip install pandas selenium openpyxl

**Arquivo:**
- automacao_cadastro_servidores.py

