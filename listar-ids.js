const axios = require('axios');
const fs = require('fs');
const path = require('path');

const tokenFile = path.join(__dirname, "tokens.json");
const tokens = JSON.parse(fs.readFileSync(tokenFile, "utf8"));

async function descobrirSede() {
    try {
        console.log("==================================================");
        console.log("🎯 BUSCANDO O PEDIDO 40566 (O DO SEU PRINT)");
        console.log("==================================================");
        
        // Busca exatamente o pedido 40566 onde você preencheu a SEDE
        const resp = await axios.get("https://www.bling.com.br/Api/v3/pedidos/vendas?numero=40566", {
            headers: { Authorization: `Bearer ${tokens.access_token}` }
        });
        
        const pedido = resp.data.data[0];
        
        if (pedido) {
            console.log(`✅ Pedido 40566 encontrado!`);
            console.log(`➡️  ID DO VENDEDOR (SITE): ${pedido.vendedor ? pedido.vendedor.id : 'Nenhum'}`);
            console.log(`➡️  ID DA LOJA (SEDE): ${pedido.loja ? pedido.loja.id : 'Nenhuma'}`);
        } else {
            console.log("❌ Pedido 40566 não encontrado.");
        }
        console.log("==================================================");

    } catch (e) {
         console.log("❌ Erro ao ler pedido:", e.response ? JSON.stringify(e.response.data) : e.message);
    }
}

descobrirSede();