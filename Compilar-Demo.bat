@echo off
rem Recompila a demonstracao da tela de Apontamentos (nao precisa do Visual Studio).
rem A lista de fontes e opcoes esta em build-demo.rsp. TemaMecano.vb, UiMecano.vb e
rem Colaborador.vb sao lidos direto de ..\TO-DO_Entrega-Cadastro-Colaboradores (sem copia).
cd /d "%~dp0"
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\vbc.exe" @build-demo.rsp
if errorlevel 1 (
    echo.
    echo ERRO na compilacao.
    pause
) else (
    echo Compilado com sucesso: GM-Apontamentos-Demo.exe
)
