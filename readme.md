## 📊 CRM LQL SOLUÇÕES - Simulador e Gestão

Este projeto é um Web App desenvolvido em Google Apps Script (GAS), integrado com o Google Sheets, destinado à simulação de empréstimos consignados e gestão de comissionamento para representantes bancários da LQL Soluções.

O sistema substitui ferramentas legadas (VBA/Access) por uma solução moderna, baseada na nuvem, acessível via browser e com controlo de perfis em tempo real.

## 🚀 Funcionalidades Principais

Autenticação Segura: Acesso restrito via "Chave J", validando promotores ativos na base de dados.

Painel Home (Dashboard): Boas-vindas personalizadas e acompanhamento de Meta Mensal com barra de progresso visual.

Simulador Dinâmico: \* Cálculo de Parcela com base no Valor do Empréstimo.

Cálculo de Margem Total com base no Valor da Parcela.

Seleção de tabelas e convénios carregados dinamicamente da folha de cálculo.

Tabelas de Comissão Inteligentes: Exibição automática das taxas de comissão filtradas pelo perfil do utilizador logado (GOLD, BLACK, TOP, GESTOR, etc.).

Registo de Auditoria (Logs): Gravação automática de todas as simulações e acessos numa aba de logs para controlo do administrador.

Interface Responsiva: Desenvolvida com Tailwind CSS, adaptando-se a telemóveis, tablets e computadores.

## 🛠️ Tecnologias Utilizadas

Backend: Google Apps Script (Javascript V8).

Frontend: HTML5, CSS3 (Tailwind CSS), JavaScript.

Base de Dados: Google Sheets (Folhas de Cálculo Google).

Iconografia: Font Awesome 6.

CDN: Google Drive para armazenamento de ativos (Logótipo).

## 📋 Estrutura da Folha de Cálculo (Database)

Para o correto funcionamento, a folha de cálculo deve conter as seguintes abas:

Promotores: Cadastro de utilizadores.

Colunas: CHAVE J, NOME, PERFIL, META, SITUAÇÃO.

bdComissao: Base de taxas de comissão.

Colunas: GRUPO, DESCRIÇÃO, Taxa Ini, Taxa Fin, Prazo Ini, Prazo Fin, [PERFIS...].

Simulador: Tabelas de fatores para o simulador.

Colunas: ID, TABELA, PRAZO, FATOR, BANCO.

Logs: Registo de atividades (gerada automaticamente pelo sistema).

## ⚙️ Configuração e Instalação

Crie uma nova Folha de Cálculo no Google Drive.

Aceda a Extensões > Apps Script.

Copie o conteúdo de code.gs para o editor de script.

Crie um novo ficheiro HTML com o nome Index e copie o conteúdo de Index.html.

Substitua o ID da folha de cálculo na função getSpreadsheet() no code.gs.

Clique em Implementar > Nova implementação e selecione "Aplicação Web".

Configure para "Executar como: Eu" e "Quem tem acesso: Qualquer pessoa".

## 👨‍💻 Desenvolvedor

Desenvolvido por um Analista de Desenvolvimento de Sistemas focado em soluções de alta performance no ecossistema Google.

"Tudo o que fizerem, façam de todo o coração, como para o Senhor." - Colossenses 3:23
