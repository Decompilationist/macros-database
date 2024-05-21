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
