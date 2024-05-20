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
    if (username === 'Control Tower' || username === 'Control Tower Formatar') {
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

app.use('/macro', authenticate);

app.get('/macro', (req, res) => {
    // Retorna a macro VBA correspondente ao usuário
    res.type('text/plain');
    res.send(req.macroVBA);
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(port, () => {
    console.log(`Servidor rodando em http://localhost:${port}`);
});
