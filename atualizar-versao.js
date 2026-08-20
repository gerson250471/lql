const fs = require('fs');

// 1. Lê a versão atualizada do package.json
const pacote = require('./package.json');
const novaVersao = pacote.version;

// 2. Lê o ficheiro Index.html
const caminhoHtml = './Index.html';
let html = fs.readFileSync(caminhoHtml, 'utf-8');

// 3. Procura a versão no formato exato do seu HTML (ex: "V: 1.4.1" ou "V:1.4.1")
// O \s* permite que haja zero ou mais espaços entre os dois pontos e os números
const regexVersao = /V:\s*\d+\.\d+\.\d+/g;
html = html.replace(regexVersao, `V: ${novaVersao}`);

// 4. Guarda o ficheiro Index.html atualizado
fs.writeFileSync(caminhoHtml, html);

console.log(`✔️ Sucesso! O Index.html foi atualizado automaticamente para a versão V: ${novaVersao}`);