# Sistema de Liberação de Motoristas

Sistema desktop desenvolvido em Python para automação do processo de liberação de motoristas, monitoramento de planilhas e envio automático de e-mails.

O projeto foi criado para centralizar e agilizar o fluxo operacional de liberações, reduzindo tarefas manuais e aumentando a confiabilidade do processo.

---

#  Funcionalidades

* Monitoramento automático de planilhas Excel
* Comparação de registros novos e já processados
* Envio automático de e-mails
* Interface gráfica moderna com PySide6
* Integração com Selenium + Edge
* Controle de logs e alertas
* Atualização contínua em loop
* Sistema preparado para execução local
* Organização modular do projeto

---

#  Tecnologias Utilizadas

## Backend

* Python 3
* Pandas
* OpenPyXL
* Selenium
* Requests

## Interface

* PySide6
* PyStray
* Pillow

## Integrações

* Outlook (pywin32)
* Microsoft Excel
* Microsoft Forms
* Microsoft Edge WebDriver

---

#  Estrutura do Projeto

```bash
Liberacao-de-Motoristas/
│
├── assets/                  # Logos e imagens do sistema
├── config/                  # Arquivos de configuração
│   ├── config.json
│   └── settings.py
│
├── core/                    # Regras de negócio
│   ├── alerts.py
│   ├── driver_setup.py
│   ├── email_handler.py
│   ├── excel_handler.py
│   └── monitor.py
│   └── processador.py
│   └── selenium_excel.py
│   └── utils.py
│
├── ui/                      # Interface gráfica
│   └── main_ui.py
│
├── main.py                  # Arquivo principal
├── requirements.txt
└── .gitignore
```

---

# Como Executar o Projeto

## 1. Clonar o repositório

```bash
git clone <url-do-repositorio>
cd Liberacao-de-Motoristas
```

---

## 2. Criar ambiente virtual

### Windows

```bash
python -m venv venv
venv\Scripts\activate
```

---

## 3. Instalar dependências

```bash
pip install -r requirements.txt
```

---

## 4. Configurar o projeto

Edite o arquivo:

```bash
config/config.json
```

Defina:

* Caminhos das planilhas
* URL do Excel Online
* Destinatários de e-mail
* Caminho do EdgeDriver
* Configurações de logging

---

## 5. Executar

```bash
python main.py
```

---

#  Funcionamento do Sistema

O sistema realiza:

1. Monitoramento da planilha principal
2. Verificação de novos registros
3. Comparação com planilha auxiliar
4. Identificação de liberações pendentes
5. Envio automático de e-mails
6. Registro das operações em log

---

#  Interface

A aplicação possui interface gráfica desenvolvida com PySide6 contendo:

* Início e pausa do monitoramento
* Status do sistema
* Histórico de liberações
* Busca de registros
* Controle operacional

---

#  Dependências Principais

```txt
pandas
openpyxl
PySide6
selenium
requests
pywin32
pystray
pillow
```

---

#  Observações

* O projeto utiliza Microsoft Outlook instalado localmente para envio de e-mails.
* O EdgeDriver deve ser compatível com a versão do Microsoft Edge instalada.
* O sistema foi desenvolvido para ambiente Windows.

---

#  Objetivo do Projeto

O objetivo do sistema é automatizar e centralizar o processo operacional de liberação de motoristas, reduzindo tempo de execução, erros manuais e dependência de controles operacionais descentralizados.

Além da automação operacional, o projeto também serve como base para futuras evoluções envolvendo centralização de dados, dashboards e escalabilidade corporativa.

---

# 👨‍💻 Autor

Desenvolvido por Arthur Mendes Sacramento.
