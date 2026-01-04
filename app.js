const fs = require('fs');
const pdf = require('pdf-parse'); // AQUI ESTAVA O ERRO. Agora está corrigido.
const path = require('path');

// Função para limpar e encontrar os dados
function extrairDados(texto) {
    // 1. Unidade Consumidora (Pega os números logo após a frase)
    // O regex procura: frase "Unidade Consumidora", quebra de linha ou espaço, e os números
    const ucMatch = texto.match(/Unidade Consumidora\s*\n*\s*(\d+)/i);

    // 2. Valor (Pega o valor monetário após "Total a pagar")
    const valorMatch = texto.match(/Total a pagar\s*\n*\s*R\$\s*([\d.,]+)/i);

    // 3. Vencimento (Pega a data após "Vencimento")
    const vencimentoMatch = texto.match(/Vencimento\s*\n*\s*(\d{2}\/\d{2}\/\d{4})/i);

    return {
        uc: ucMatch ? ucMatch[1] : 'NÃO ENCONTRADO',
        valor: valorMatch ? valorMatch[1] : 'NÃO ENCONTRADO',
        vencimento: vencimentoMatch ? vencimentoMatch[1] : 'NÃO ENCONTRADO'
    };
}

async function iniciar() {
    // Lista os arquivos da pasta atual
    const arquivos = fs.readdirSync(__dirname).filter(f => f.toLowerCase().endsWith('.pdf'));

    console.log(`Encontrados ${arquivos.length} arquivos PDF.\n`);

    for (const arquivo of arquivos) {
        try {
            const caminhoCompleto = path.join(__dirname, arquivo);
            const buffer = fs.readFileSync(caminhoCompleto);
            
            // AQUI É ONDE O ERRO ACONTECIA. Com a correção lá em cima, vai funcionar.
            const data = await pdf(buffer);
            
            const dados = extrairDados(data.text);

            console.log(`📄 Arquivo: ${arquivo}`);
            console.log(`   📌 UC: ${dados.uc}`);
            console.log(`   📅 Vencimento: ${dados.vencimento}`);
            console.log(`   💰 Valor: R$ ${dados.valor}`);
            console.log('-----------------------------------------');

        } catch (erro) {
            console.log(`❌ Erro no arquivo ${arquivo}:`);
            console.log(`   Mensagem: ${erro.message}`);
            // Dica de debug: se der erro aqui, mostra o tipo da variável pdf
            if (erro.message.includes("is not a function")) {
                console.log(`   ALERTA: A importação da biblioteca falhou. Tipo atual: ${typeof pdf}`);
            }
        }
    }
}

iniciar();