# 📊 Macros Database API

Este projeto é uma API em Node.js utilizando Express, que fornece acesso a macros VBA armazenadas em um arquivo/servidor `.env`. A autenticação básica é utilizada para proteger o acesso às macros.

## ✨ Funcionalidades

- 🔐 **Autenticação**: Utiliza autenticação básica com nome de usuário e senha.
- 📥 **Recuperação de Macros VBA**: Busca macros VBA específicas para usuários autenticados.
- 📋 **Listagem de Títulos de Macros**: Retorna todos os títulos das macros VBA disponíveis no arquivo `.env`.

## 📋 Pré-requisitos

- Node.js
- npm (Node Package Manager)
- Arquivo `.env` configurado com as macros VBA e senha de autenticação

## ⚙️ Configuração do Ambiente

1. Clone o repositório:

   ```bash
   git clone https://github.com/seu-usuario/macros-database.git
   cd macros-database
   ```

2. Instale as dependências:

   ```bash
   npm install
   ```

3. Crie um arquivo .env na raiz do projeto com o seguinte formato:

   ```bash
   PASS=sua_senha
   MACRO_VBA_CONTROL_TOWER_1=Sub Macro1() ' código VBA aqui
   End Sub
   MACRO_VBA_CONTROL_TOWER_2=Sub Macro2() ' código VBA aqui
   End Sub
   ```
## 🚀 Uso

1. Inicie o servidor:

   ```bash
   npm start
