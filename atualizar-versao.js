const fs = require('fs');

// 1. Lê a versão atualizada do package.json
const pacote = require('./package.json');
const novaVersao = pacote.version;

// 2. Lê o ficheiro Index.html
const caminhoHtml = './Index.html';
let html = fs.readFileSync(caminhoHtml, 'utf-8');

// 3. Procura qualquer versão no formato vX.X.X e substitui pela nova
const regexVersao = /v\d+\.\d+\.\d+/g;
html = html.replace(regexVersao, `v${novaVersao}`);

// 4. Guarda o ficheiro Index.html atualizado
fs.writeFileSync(caminhoHtml, html);

console.log(`✔️ Sucesso! O Index.html foi atualizado automaticamente para a versão v${novaVersao}`);