const axios = require("axios");
const fs = require("fs");

// COLOQUE SEUS DADOS AQUI:
const CLIENT_ID = "0e1b32755e0a5c0d22bfaf0c6e58d7de886e45ee";
const CLIENT_SECRET = "50bdda91881b0b1981c7529677f5b99e825ffe9f7d7bfcffa422bde94f26";

// 👇 NOVO CÓDIGO FRESQUINHO (Válido por 1 minuto)
const CODIGO_DO_NAVEGADOR = "a95ae930a1d95de4932d0b9d5f81c7c1c3993be6";

const credenciais = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString("base64");

axios.post(
  "https://www.bling.com.br/Api/v3/oauth/token",
  new URLSearchParams({ grant_type: "authorization_code", code: CODIGO_DO_NAVEGADOR }),
  { headers: { Authorization: `Basic ${credenciais}`, "Content-Type": "application/x-www-form-urlencoded" } }
).then(res => {
  const tokens = {
    access_token: res.data.access_token,
    refresh_token: res.data.refresh_token,
    expires_at: Date.now() + res.data.expires_in * 1000,
  };
  fs.writeFileSync("tokens.json", JSON.stringify(tokens, null, 2));
  console.log("✅ SUCESSO! O arquivo tokens.json foi criado com as novas permissões do WMS!");
}).catch(err => console.error("❌ Erro:", err.response ? err.response.data : err.message));