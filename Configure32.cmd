@echo off
if exist "Configuration.xml" del "Configuration.xml" /f /q
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
 echo   ^<Add SourcePath=".\Office\" OfficeClientEdition="32" Channel="Current" Version="16.0.17628.20110"^> >> Configuration.xml
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
 echo   ^<Add SourcePath=".\Office\" OfficeClientEdition="32" Channel="Current" Version="16.0.17628.20110"^> >> Configuration.xml
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
 echo   ^<Add SourcePath=".\Office\" OfficeClientEdition="32" Channel="Current" Version="16.0.17628.20110"^> >> Configuration.xml
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
 echo   ^<Add SourcePath=".\Office\" OfficeClientEdition="32" Channel="PepetualVL2024" Version="16.0.17531.20154"^> >> Configuration.xml
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
if exist "Configure64.cmd" del "Configure64.cmd" /f /q
if exist "Configure32.cmd" del "Configure32.cmd" /f /q
exit