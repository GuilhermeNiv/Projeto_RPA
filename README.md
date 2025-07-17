# ✉️ Projeto de Manipulação de Planilhas e Envio de E-mail

Este projeto automatiza a leitura de três planilhas Excel, consolida os dados em um único arquivo e envia via e-mail para o remetente de sua escolha. Meu objetivo com este projeto é mostrar de maneira simples a possibilidade de automatizar tarefas repetitivas, neste caso montando planilhas consolidadas e enviando-as via e-mail automaticamente.

---

## 💡 Funcionalidades

- Análise e manipulação de dados em planilhas
- Disparo de e-mail com o arquivo consolidado  

---

## 🛠️ Tecnologias Utilizadas

- Python 3.7 ou superior
- Pacotes Python:
  - pandas
  - openpyxl
  - python-dotenv
- Conta Gmail com "Acesso a apps menos seguros" ativado (ou token de app, caso use 2FA)

---

## 📁 Componentes do Projeto

Projeto_RPA/
│
├── Planilhas/
│ ├── vendas_01.xlsx
│ ├── vendas_02.xlsx
│ └── vendas_03.xlsx
│
├── .env
├── automacao.py

---

1. Instale as dependências:  

   pip install pandas openpyxl python-dotenv
   
2. Coloque as planilhas vendas_01.xlsx, vendas_02.xlsx e vendas_03.xlsx dentro da pasta Planilhas.

3. Crie um arquivo .env no seu projeto com este conteúdo:

EMAIL_REMETENTE=seuemail@gmail.com
SENHA_REMETENTE=sua_senha_ou_token_de_app
EMAIL_DESTINATARIO=email_destino@gmail.com

Obs: A senha do remetente a qual me refiro acima é a gerada pelo seu provedor de domínio, geralmente composta por 16 caracteres aleatórios

---

## 🧪 Como testar

Execute o script automacao.py:

python automacao.py

O script irá:

- Validar as variáveis do .env

- Carregar e consolidar os dados das planilhas

- Gerar um arquivo Excel resumo com os dados consolidados

- Enviar um e-mail com o arquivo resumo em anexo

Obs: As planilhas devem estar nomeadas exatamente como vendas_01.xlsx, vendas_02.xlsx e vendas_03.xlsx, para o código funcionar.

---

👨‍💻 Autor:
Guilherme Laureano | 
Disponível para contratação | Uberlândia - MG
