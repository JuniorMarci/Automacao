@echo off
chcp 65001

set "sourceFile=Correções (V2)\Anaálise de Metas V2.pbix"
set "destinationFolder=Hist"
set "destinationFile=Anaálise de Metas -- Somente Leitura.pbix"

REM Copia o arquivo para o destino
copy "%sourceFile%" "%destinationFolder%\%destinationFile%"

REM Abre o arquivo
start "" "%destinationFolder%\%destinationFile%"

REM Aguarda até que o Power BI seja fechado
:waitForClose
timeout /t 2 /nobreak >nul
tasklist /fi "imagename eq PBIDesktop.exe" | find /i "PBIDesktop.exe" >nul
if errorlevel 1 (
    REM Deleta o arquivo
    del "%destinationFolder%\%destinationFile%"
    exit
)
goto :waitForClose
