# 🐍 Ferramentas de automação em Python 

## 🇧🇷 Português (idioma principal || main language)

Este repositório contém uma coleção de scripts Python para automações como extração de relatórios e interações com um banco de dados PostgreSQL local.

### 🚀 Funcionalidades

- Extração automatizada de relatórios
- Operações com banco de dados PostgreSQL (leitura/gravação/atualização)
- Estrutura de código modular e escalável

### 🛠 Guia de Instalação

1. **Clone o repositório**

```bash
git clone https://github.com/pedro-alpack/alpack-ics-dataExtract.git
cd alpack-ics-dataExtract
```

2. **Crie e ative um ambiente virtual**

```bash
python -m venv venv
source venv/bin/activate  # No Windows use `venv\Scripts\activate`
```

3. **Instale as dependências**

```bash
pip install -r requirements.txt
```

4. **Configure o arquivo `.env`**

Crie um arquivo `.env` no diretório raiz e adicione suas credenciais do PostgreSQL:

```
DB_HOST=SEU NOME DO HOST
DB_NAME=SEU NOME DO DB
DB_USER=SEU USUÁRIO
DB_PASSWORD=SUA SENHA
DB_PORT=NÚMERO DA PORTA
```

5. **Execute o script desejado**

```bash
python scripts/seu_script.py
```

### 🛠 Configuração das Tabelas no Banco de Dados

Para criar as tabelas necessárias no seu banco de dados PostgreSQL, você pode rodar as seguintes queries SQL:

**Tabela 1: `ligacoes`**

```sql
CREATE TABLE ligacoes (
    id SERIAL PRIMARY KEY,
    data DATE NOT NULL,
    horario TIME WITHOUT TIME ZONE NOT NULL,
    ramal INTEGER NOT NULL,
    duracao TIME WITHOUT TIME ZONE NOT NULL,
    ligid TEXT NOT NULL
);
```

**Tabela 2: `vendas`**

```sql
CREATE TABLE vendas (
    id SERIAL PRIMARY KEY,
    data DATE NOT NULL,
    vendedor TEXT NOT NULL,
    valor_vendido NUMERIC(10,2) NOT NULL,
    statusPedido TEXT NOT NULL,
    cliente TEXT NOT NULL,
    regiao TEXT NOT NULL,
    vendaID NUMERIC NOT NULL
);
```
## Nota

Esse projeto foi desenvolvido em cima de aplicações específicas utilizadas na empresa Alpack do Brasil (https://www.alpack.com.br/). O repositório público no github tem como único objetivo compartilhar o trabalho desenvolvido por nosso desenvolvedor ([@pm-ramoss](https://github.com/pm-ramoss)).

---

# 🐍 Python Automation Toolkit

## 🇺🇸 English

This repository contains a collection of Python automation scripts for data extraction, report generation, and interaction with a local PostgreSQL database.

### 🚀 Features

- Automated report extraction
- PostgreSQL database operations (read/write/update)
- Modular and scalable code structure

### 🛠 Installation Guide

1. **Clone the repository**

```bash
git clone https://github.com/pedro-alpack/alpack-ics-dataExtract.git
cd alpack-ics-dataExtract
```

2. **Create a virtual environment and activate it**

```bash
python -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
```

3. **Install dependencies**

```bash
pip install -r requirements.txt
```

4. **Set up your `.env` file**

Create a `.env` file in the root directory and add your PostgreSQL credentials:

```
DB_HOST=YOUR HOSTNAME
DB_NAME=YOUR DB NAME
DB_USER=YOUR USER
DB_PASSWORD=YOUR PASSWORD
DB_PORT=YOUR PORT NUMBER
```

5. **Run your desired script**

```bash
python scripts/your_script.py
```

### 🛠 Database Tables Setup

To create the necessary tables in your PostgreSQL database, you can run the following SQL queries:

**Table 1: `ligacoes`**

```sql
CREATE TABLE ligacoes (
    id SERIAL PRIMARY KEY,
    data DATE NOT NULL,
    horario TIME WITHOUT TIME ZONE NOT NULL,
    ramal INTEGER NOT NULL,
    duracao TIME WITHOUT TIME ZONE NOT NULL,
    ligid TEXT NOT NULL
);
```

**Table 2: `vendas`**

```sql
CREATE TABLE vendas (
    id SERIAL PRIMARY KEY,
    data DATE NOT NULL,
    vendedor TEXT NOT NULL,
    valor_vendido NUMERIC(10,2) NOT NULL,
    statusPedido TEXT NOT NULL,
    cliente TEXT NOT NULL,
    regiao TEXT NOT NULL,
    vendaID NUMERIC NOT NULL
);
```
## Note

This project was developed based on specific applications used by the company Alpack do Brasil (https://www.alpack.com.br/). The public repository on GitHub is solely intended to share the work developed by our developer ([@pm-ramoss](https://github.com/pm-ramoss)).

---