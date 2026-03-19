@echo off
title Servidor - Sistema AJ Modas
color 0B
echo ===================================================
echo         INICIANDO SISTEMA AJ MODAS...
echo ===================================================
echo.
echo ATENCAO: Nao feche esta tela preta! 
echo E ela que mantem o Robo do WhatsApp e o Checkout funcionando.
echo Se quiser desligar o sistema, basta fechar esta janela.
echo.
echo Abrindo o painel no navegador...
timeout /t 3 >nul
start http://localhost:3000/checkout.html
node index.js
pause