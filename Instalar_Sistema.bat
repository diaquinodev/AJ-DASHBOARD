@echo off
title Instalador - Sistema AJ Modas
color 0A
echo ===================================================
echo      INSTALADOR AUTOMATICO - SISTEMA AJ MODAS
echo ===================================================
echo.
echo Verificando se o Node.js esta instalado no computador...
node -v >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Node.js NAO ENCONTRADO!
    echo O sistema vai abrir o site oficial agora.
    echo Baixe a versao "LTS", instale (clicando em Next, Next, Install) e rode este instalador novamente.
    timeout /t 5 >nul
    start https://nodejs.org/
    pause
    exit
)
echo [OK] Node.js detectado!
echo.
echo Baixando o motor do WhatsApp e as bibliotecas do sistema...
echo Isso pode demorar alguns minutos dependendo da internet. Aguarde...
echo.
npm install express cors axios node-cron whatsapp-web.js qrcode-terminal xlsx multer pdf-parse
echo.
echo ===================================================
echo   INSTALACAO CONCLUIDA COM SUCESSO, DIEGO!
echo   Voce ja pode fechar esta janela e iniciar o sistema.
echo ===================================================
pause