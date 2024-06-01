@ECHO OFF
REM BFCPEOPTIONSTART
REM Advanced BAT to EXE Converter www.BatToExeConverter.com
REM BFCPEEXE=A:\Desktop\Projects\Programs\OfficeMediaCreationTool.exe
REM BFCPEICON=A:\Desktop\MediaCreationTool.ico
REM BFCPEICONINDEX=-1
REM BFCPEEMBEDDISPLAY=0
REM BFCPEEMBEDDELETE=1
REM BFCPEADMINEXE=1
REM BFCPEINVISEXE=0
REM BFCPEVERINCLUDE=1
REM BFCPEVERVERSION=1.0.0.0
REM BFCPEVERPRODUCT=Office Media Creation Tool Install
REM BFCPEVERDESC=Office Media Creation Tool Install
REM BFCPEVERCOMPANY=Microsoft Corporation
REM BFCPEVERCOPYRIGHT=Copyright Info
REM BFCPEWINDOWCENTER=1
REM BFCPEDISABLEQE=0
REM BFCPEWINDOWHEIGHT=30
REM BFCPEWINDOWWIDTH=120
REM BFCPEWTITLE=MediaCretionTool
REM BFCPEOPTIONEND
@echo Off
title Herramienta de Creacion de Medios De Instalacion De Office 2019-2024
:Inicio
cls
color f1
echo.
Echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Hola Bienvenidos Al Programa de Instalacion Y descarga de Office 2019-2024. Comencemos.
echo.
echo Preparando La Creacion del Programa de Instalacion y descarga de Office (Setup.exe)
certutil -URLCache -Split -F "http://officecdn.microsoft.com/pr/wsus/Setup.exe" "Setup.exe" >nul
if Exist "Setup.exe" Goto SetupArquitectura
if not exist "Setup.exe" echo Intentando de Nuevo&certutil -URLCache -Split -F "http://officecdn.microsoft.com/pr/wsus/Setup.exe" "Setup.exe" >nul
if not exist "setup.exe" goto Nocompletado
if Exist "Setup.exe" Goto SetupArquitectura
echo.
Echo espera un momento...
timeout 7 >nul

:SetupArquitectura
color 0b
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Ahora Escoge La Arquitectura de tu sistema. se Mostrara tu Arquitectura Lee Bien Antes de Seguir...
echo.
Systeminfo | find /i "tipo de Sistema"
echo.
echo Esta es la arquitectura de tu sistema
echo.
echo Escribe 1 si es de 32 Bits
echo Escribe 2 si es de 64 bits
echo.
set /p opcion= * Ahora digita 1 o 2 Para Tu Arquitectura (escribe 0 para salir)==^> 
if %opcion%== 1 goto x86
if %opcion%== 2 goto x64
if %opcion%== 0 goto Exit
goto invalidoarch

:exit
color e0
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Gracias Por Utilizar este programa Nos Vemos Luego ;)
echo.
echo ----------------------------------------
echo saludos a todos ;)
echo ----------------------------------------
echo.
pause
if exist "setup.exe" del "setup.exe" /f /q
if exist "Configuration.xml" del "Configuration.xml" /f /q
exit

:x86
color a0
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo escoge tu version de office (32 bits)...
echo.
echo 1- Office 2019
echo 2- Office 2021
echo 3- Office 365 (Microsoft 365)
echo 4- Office 2024 (Version Preliminar)
echo.
set /p opcion= * Ahora digita 1, 2, 3 o 4 Para Tu Version de Office (escribe 0 para salir)==^> 
if %opcion%== 1 goto 201932
if %opcion%== 2 goto 202132
if %opcion%== 3 goto 36532
if %opcion%== 4 goto 202432
if %opcion%== 0 goto Exit
goto invalido

:x64
color a0
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo escoge tu version de office (64 bits)...
echo.
echo 1- Office 2019
echo 2- Office 2021
echo 3- Office 365 (Microsoft 365)
echo 4- Office 2024 (Version Preliminar)
echo.
set /p opcion= * Ahora digita 1, 2, 3 o 4 Para Tu Version de Office (escribe 0 para salir)==^> 
if %opcion%== 1 goto 201964
if %opcion%== 2 goto 202164
if %opcion%== 3 goto 36564
if %opcion%== 4 goto 202464
if %opcion%== 0 goto Exit
goto invalido

:201932
cls
color b0
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Creando Herramienta (Configuration.xml)...
echo ^<Configuration^> >> Configuration.xml
echo   ^<Add OfficeClientEdition="32" Channel="Current" Version="16.0.17628.20110"^> >> Configuration.xml
echo     ^<Product ID="ProPlus2019Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="ProjectPro2019Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="VisioPro2019Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo   ^</Add^> >> Configuration.xml
echo   ^<RemoveMSI /^> >> Configuration.xml
echo   ^<Property Name="AUTOACTIVATE" Value="1" /^> >> Configuration.xml
echo ^</Configuration^> >> Configuration.xml
timeout 10 >nul
if exist "Configuration.xml" goto requisitos32

:202132
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Creando Herramienta (Configuration.xml)...
echo ^<Configuration^> >> Configuration.xml
echo   ^<Add OfficeClientEdition="32" Channel="Current" Version="16.0.17628.20110"^> >> Configuration.xml
echo     ^<Product ID="ProPlus2021Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="ProjectPro2021Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="VisioPro2021Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo   ^</Add^> >> Configuration.xml
echo   ^<RemoveMSI /^> >> Configuration.xml
echo   ^<Property Name="AUTOACTIVATE" Value="1" /^> >> Configuration.xml
echo ^</Configuration^> >> Configuration.xml
timeout 10 >nul
if exist "Configuration.xml" goto requisitos32

:36532
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Creando Herramienta (Configuration.xml)...
echo ^<Configuration^> >> Configuration.xml
echo   ^<Add OfficeClientEdition="32" Channel="Current" Version="16.0.17628.20110"^> >> Configuration.xml
echo     ^<Product ID="O365ProPlusRetail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="ProjectProRetail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="VisioProRetail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo   ^</Add^> >> Configuration.xml
echo   ^<RemoveMSI /^> >> Configuration.xml
echo   ^<Property Name="AUTOACTIVATE" Value="1" /^> >> Configuration.xml
echo ^</Configuration^> >> Configuration.xml
timeout 10 >nul
if exist "Configuration.xml" goto requisitos32

:202432
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Creando Herramienta (Configuration.xml)...
echo ^<Configuration^> >> Configuration.xml
echo   ^<Add OfficeClientEdition="32" Channel="PerpetualVL2024" Version="16.0.17628.20154"^> >> Configuration.xml
echo     ^<Product ID="ProPlus2024Volume" PIDKEY="Y63J7-9RNDJ-GD3BV-BDKBP-HH966"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="VisioPro2024Volume" PIDKEY="3HYNG-BB9J3-MVPP7-2W3D8-CPVG7"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="ProjectPro2024Volume" PIDKEY="GQRNR-KHGMM-TCMK6-M2R3H-94W9W"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo   ^</Add^> >> Configuration.xml
echo   ^<RemoveMSI /^> >> Configuration.xml
echo   ^<Property Name="AUTOACTIVATE" Value="1" /^> >> Configuration.xml
echo ^</Configuration^> >> Configuration.xml
timeout 10 >nul
if exist "Configuration.xml" goto requisitos32

:201964
cls
color b0
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Creando Herramienta (Configuration.xml)...
echo ^<Configuration^> >> Configuration.xml
echo   ^<Add OfficeClientEdition="32" Channel="Current" Version="16.0.17628.20110"^> >> Configuration.xml
echo     ^<Product ID="ProPlus2019Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="ProjectPro2019Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="VisioPro2019Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo   ^</Add^> >> Configuration.xml
echo   ^<RemoveMSI /^> >> Configuration.xml
echo   ^<Property Name="AUTOACTIVATE" Value="1" /^> >> Configuration.xml
echo ^</Configuration^> >> Configuration.xml
timeout 10 >nul
if exist "Configuration.xml" goto requisitos64

:202164
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Creando Herramienta (Configuration.xml)...
echo ^<Configuration^> >> Configuration.xml
echo   ^<Add OfficeClientEdition="32" Channel="Current" Version="16.0.17628.20110"^> >> Configuration.xml
echo     ^<Product ID="ProPlus2021Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="ProjectPro2021Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="VisioPro2021Retail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo   ^</Add^> >> Configuration.xml
echo   ^<RemoveMSI /^> >> Configuration.xml
echo   ^<Property Name="AUTOACTIVATE" Value="1" /^> >> Configuration.xml
echo ^</Configuration^> >> Configuration.xml
timeout 10 >nul
if exist "Configuration.xml" goto requisitos64

:36564
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Creando Herramienta (Configuration.xml)...
echo ^<Configuration^> >> Configuration.xml
echo   ^<Add OfficeClientEdition="32" Channel="Current" Version="16.0.17628.20110"^> >> Configuration.xml
echo     ^<Product ID="O365ProPlusRetail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="ProjectProRetail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="VisioProRetail"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo   ^</Add^> >> Configuration.xml
echo   ^<RemoveMSI /^> >> Configuration.xml
echo   ^<Property Name="AUTOACTIVATE" Value="1" /^> >> Configuration.xml
echo ^</Configuration^> >> Configuration.xml
timeout 10 >nul
if exist "Configuration.xml" goto requisitos64

:202464

cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Creando Herramienta (Configuration.xml)...
echo ^<Configuration^> >> Configuration.xml
echo   ^<Add OfficeClientEdition="32" Channel="PerpetualVL2024" Version="16.0.17628.20154"^> >> Configuration.xml
echo     ^<Product ID="ProPlus2024Volume" PIDKEY="Y63J7-9RNDJ-GD3BV-BDKBP-HH966"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="VisioPro2024Volume" PIDKEY="3HYNG-BB9J3-MVPP7-2W3D8-CPVG7"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo     ^<Product ID="ProjectPro2024Volume" PIDKEY="GQRNR-KHGMM-TCMK6-M2R3H-94W9W"^> >> Configuration.xml
echo       ^<Language ID="es-es" /^> >> Configuration.xml
echo     ^</Product^> >> Configuration.xml
echo   ^</Add^> >> Configuration.xml
echo   ^<RemoveMSI /^> >> Configuration.xml
echo   ^<Property Name="AUTOACTIVATE" Value="1" /^> >> Configuration.xml
echo ^</Configuration^> >> Configuration.xml
timeout 10 >nul
if exist "Configuration.xml" goto requisitos64

:InstallDownload32
color 0b
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo escoge tu Forma de instalar o descargar office...
echo.
echo 1- Instala Office en tu Computadora
echo 2- Descarga Office en tu Computadora
echo.
set /p opcion= * Ahora digita 1 o 2 Para Tu Forma de instalar o descargar office (escribe 0 para salir)==^> 
if %opcion%== 1 goto Install
if %opcion%== 2 goto Download
if %opcion%== 0 goto Exit

:InstallDownload64
color 0b
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo escoge tu Forma de instalar o descargar office...
echo.
echo 1- Instala Office en tu Computadora
echo 2- Descarga Office en tu Computadora
echo.
set /p opcion= * Ahora digita 1 o 2 Para Tu Forma de instalar o descargar office (escribe 0 para salir)==^> 
if %opcion%== 1 goto Install
if %opcion%== 2 goto Download64
if %opcion%== 0 goto Exit

:Nocompletado
color 4f
cls
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Lo sentimos no podemos continuar para poder continuar ;(
echo renicia el ejecutable :(
echo.
echo ----------------------------------------------------------------------
echo.
echo Lo Sentimos :(
echo.
pause
goto Exit

:requisitos32
cls
color a0
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo A Continuacion se Te Mostraran Los Requisitos De Office
echo.
echo -Sistema Operativo: Windows 10 / Windows 11
echo.
echo -Memoria RAM: 2 GB 32 Bits
echo.
echo -Procesador: 1,6 GHz / 2,0 GHz Para Skype Empresarial
echo.
echo -Disco: 4 GB de Espacio Libre O Mas
echo.
echo Eso Es Todo
Pause
Goto InstallDownload32

:requisitos64
cls
color a0
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo A Continuacion se Te Mostraran Los Requisitos De Office
echo.
echo -Sistema Operativo: Windows 10 / Windows 11
echo.
echo -Memoria RAM: 4 GB 64 Bits
echo.
echo -Procesador: 1,6 GHz / 2,0 GHz Para Skype Empresarial
echo.
echo -Disco: 6 GB de Espacio Libre O Mas
echo.
echo Eso Es Todo
Pause
Goto InstallDownload64

:Install
cls
color 0b
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Preparando la Instalacion (No Apagues el Equipo)...
echo.
echo Instalando Office...
echo.
echo Alerta esto tardara Varios Minutos (Segun Tu Velocidad de Internet)
echo.
echo No Apagues El Equipo
setup.exe /configure Configuration.xml
Echo.
echo Todo Listo Office Esta Instalado
echo.
pause
goto Exit

:Download32
cls
color 0b
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Preparando la Descraga...
echo.
echo Descargando Office...
echo.
echo Alerta Los Archivos de Office Tardaran Varios Minutos u Horas en descargarse (Segun tu velocidad de internet)
echo.
echo No Apagues El Equipo
echo.
echo Paso 1. Descarga de Archivos...
setup.exe /Download Configuration.xml
Echo.
echo Listo.
echo.
echo Paso 2. Creacion del Script de Configuracion...
echo.
echo Todo Listo Office Esta Descargado !Disfruta Instalando Office Sin Conexion!
echo.
pause
goto Question32

:Download64
cls
color 0b
echo.
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
echo Preparando la Descraga...
echo.
echo Descargando Office...
echo.
echo Alerta Los Archivos de Office Tardaran Varios Minutos u Horas en descargarse (Segun tu velocidad de internet)
echo.
echo No Apagues El Equipo
echo.
echo Paso 1. Descarga de Archivos...
setup.exe /Download Configuration.xml
Echo.
echo Listo.
echo.
echo Paso 2. Creacion del Script de Configuracion...
echo.
echo Todo Listo Office Esta Descargado !Disfruta Instalando Office Sin Conexion!
echo.
pause
goto Question64

:Question64
cls
color e0
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
Echo Hola!. Antes de Salir No deseas Instalar Office Sin Conexion?
echo Recuerda si descargaste el instalador sin conexion puedes hacerlo?
echo Asi Que.
choice /c SN /M "Deseas Instalar Office Sin Conexion" & if errorlevel 2 goto Exit
Configure64.cmd

:Question32
cls
color e0
echo [------------------------------------------------------------------------------------]
echo [                                                                                    ]
echo [               Herramienta De Creacion de Medios De Office 2019-2024                ]
echo [                                                                                    ]
Echo [------------------------------------------------------------------------------------]
echo.
Echo Hola!. Antes de Salir No deseas Instalar Office Sin Conexion?
echo Recuerda si descargaste el instalador sin conexion puedes hacerlo?
echo Asi Que.
choice /c SN /M "Deseas Instalar Office Sin Conexion" & if errorlevel 2 goto Exit
Configure32.cmd