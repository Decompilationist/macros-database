const express = require('express');
const basicAuth = require('basic-auth');
const dotenv = require('dotenv');
const path = require('path');

dotenv.config();

const app = express();
const port = 3000;

const PASS = process.env.PASS;

// Função para buscar a macro VBA com base no nome de usuário
function getMacroVBA(username) {
    // Verifica se o usuário é 'Control Tower' ou 'Control Tower Formatar'
    if (username === 'Control Tower' || username === 'Salvar Obrigatorio') {
        const macroKeys = Object.keys(process.env)
            .filter(key => key.startsWith(`MACRO_VBA_${username.replace(' ', '_').toUpperCase()}`))
            .sort((a, b) => {
                const numA = parseInt(a.split('_').pop());
                const numB = parseInt(b.split('_').pop());
                return numA - numB;
            });
        
        return macroKeys.map(key => process.env[key]).join('\n\n');
    }
    // Caso contrário, retorna null
    else {
        return null;
    }
}

// Função para buscar todos os títulos das macros VBA do arquivo .env e formatá-los
function getAllMacroTitles() {
    const macroTitles = new Set();
    for (const key in process.env) {
        if (key.startsWith('MACRO_VBA')) {
            const macroName = key.substring(key.indexOf("_") + 1);
            const macroTitle = macroName.split('_').slice(1).join('_'); // Removendo "VBA" do título
            const formattedTitle = macroTitle.replace(/\d+$/, '') // Remove números no final
                                              .replace(/_/g, ' ') // Substitui underscores por espaços
                                              .toLowerCase() // Transforma em minúsculas
                                              .replace(/(?:^|\s)\S/g, char => char.toUpperCase()); // Capitaliza primeira letra de cada palavra
            macroTitles.add(formattedTitle);
        }
    }
    return [...macroTitles];
}


function authenticate(req, res, next) {
    const user = basicAuth(req);
    if (user && user.pass === PASS) {
        // Verifica se a macro VBA correspondente ao usuário existe
        const macroVBA = getMacroVBA(user.name);
        if (macroVBA !== null) {
            req.macroVBA = macroVBA;
            return next();
        } else {
            return res.status(401).send('Macro não autorizado.');
        }
    } else {
        res.set('WWW-Authenticate', 'Basic realm="example"');
        return res.status(401).send('Autenticação necessária.');
    }
}

// Função para autenticar com base apenas na senha do arquivo .env
function authenticateadmin(req, res, next) {
    const credentials = basicAuth(req);
    if (credentials && credentials.pass === PASS) {
        return next();
    } else {
        res.set('WWW-Authenticate', 'Basic realm="example"');
        return res.status(401).send('Autenticação necessária.');
    }
}

app.use('/macro', authenticate);

app.get('/macro', (req, res) => {
    // Retorna a macro VBA correspondente ao usuário
    res.type('text/plain');
    res.send(req.macroVBA);
});

app.get('/macros', authenticateadmin, (req, res) => {
    // Retorna os títulos das macros VBA do arquivo .env
    res.json(getAllMacroTitles());
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(port, () => {
    console.log(`Servidor rodando em http://localhost:${port}`);
});
