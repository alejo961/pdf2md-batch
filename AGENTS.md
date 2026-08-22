# Instalación de PDF2MD por agentes de IA

Cuando el usuario solicite instalar PDF2MD desde este repositorio en Windows 10/11 de 64 bits:

1. No uses `app.py` ni instales dependencias de desarrollo como instalación final.
2. Ejecuta desde la raíz del repositorio:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
   ```

3. El script descarga únicamente el instalador de la GitHub Release oficial, verifica su SHA-256 y lo instala en `%LOCALAPPDATA%\Programs\PDF2MD`.
4. Confirma que exista `PDF2MD.lnk` en el escritorio del usuario y que `PDF2MD.exe` abra la aplicación en el navegador.
5. No solicites permisos de administrador. Si Windows muestra SmartScreen, pide al usuario que confirme la ejecución del instalador no firmado.

La instalación es local y no debe cargar documentos del usuario a internet.
