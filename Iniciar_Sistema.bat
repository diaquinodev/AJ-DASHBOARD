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

REM Libera a porta 3000 no Firewall do Windows para acesso pela rede
echo Liberando acesso pela rede (porta 3000)...
netsh advfirewall firewall delete rule name="AJ Modas - Servidor" >nul 2>&1
netsh advfirewall firewall add rule name="AJ Modas - Servidor" dir=in action=allow protocol=TCP localport=3000 >nul 2>&1
echo [OK] Porta liberada!
echo.

echo Abrindo o painel no navegador...
timeout /t 3 >nul
start http://localhost:3000/checkout.html
node index.js
pause
