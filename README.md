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


## 🔌 Endpoints

### `GET /`

Retorna a página inicial.

### `GET /macro`

Autenticação necessária: **Sim** (Nome de usuário e senha)

Retorna a macro VBA correspondente ao nome de usuário autenticado.

#### Exemplo de requisição

```bash
curl -u Control\ Tower:sua_senha http://localhost:3000/macro
```

### `GET /macros`

Autenticação necessária: **Sim** (Apenas senha)

Retorna todos os títulos das macros VBA disponíveis no arquivo `.env`.

#### Exemplo de requisição

```bash
curl -u :sua_senha http://localhost:3000/macros
```
#### Exemplo de resposta

```json
[
  "Macro1",
  "Macro2"
]
```

## 🛠️ Estrutura do Código

### Dependências

- `express`: Framework web para Node.js.
- `basic-auth`: Middleware para autenticação básica HTTP.
- `dotenv`: Carrega variáveis de ambiente de um arquivo `.env`.

### Configuração do Servidor

- Carrega as variáveis de ambiente do arquivo `.env`.
- Configura a porta do servidor para `3000`.

### Funções Principais

- **`getMacroVBA(username)`**: Busca a macro VBA para um determinado nome de usuário.
- **`getAllMacroTitles()`**: Retorna todos os títulos das macros VBA do arquivo `.env`.
- **`authenticate(req, res, next)`**: Middleware para autenticação com nome de usuário e senha.
- **`authenticateadmin(req, res, next)`**: Middleware para autenticação apenas com a senha.

### Rotas

- **`/macro`**: Protegida por `authenticate`, retorna a macro VBA do usuário autenticado.
- **`/macros`**: Protegida por `authenticateadmin`, retorna os títulos das macros VBA.

## 🤝 Contribuição

Contribuições são bem-vindas! Por favor, envie um pull request ou abra uma issue para discutir mudanças.

## 📄 Licença

Este projeto está licenciado sob a [MIT License](LICENSE).
