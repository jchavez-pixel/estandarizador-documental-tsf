# Estandarizador Documental TSF

Aplicación de procesamiento temporal de documentos PDF y DOCX, con interfaz web FastAPI y salida DOCX + PDF.

## ¿Qué hace?
- Recibe un PDF con texto seleccionable o un DOCX.
- Solicita solo tipo documental, nombre, elaborador y revisor; aprobador es opcional.
- Genera documentos aplicando un formato corporativo y entrega los resultados en ZIP.
- Emplea Microsoft Word COM en Windows (si está disponible) o LibreOffice para exportar a PDF.

> **IMPORTANTE:** Este repositorio contiene el código fuente, no un sitio que GitHub Pages pueda ejecutar. GitHub Pages no admite servidores Python. Para que la aplicación funcione cuando tu computadora está apagada, necesita un servicio de alojamiento compatible con Python y un motor de conversión PDF instalado.

## Ejecutar en Windows
Se recomienda Python 3.12.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --only-binary=:all: -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

Acceder a http://127.0.0.1:8765.

## Privacidad y seguridad
- Los documentos se procesan temporalmente en el servidor y se eliminan conforme al flujo de aplicación.
- No se agregan a GitHub automáticamente ni se almacenan en una base de datos.
- Quedan copias descargadas en el equipo del usuario y podrían existir registros temporales o de infraestructura.
- **No conectar directamente la aplicación a un dominio público de uso masivo sin controles de seguridad**, tales como límites de concurrencia y tasa, restricciones de archivos y protección de la estación de trabajo.
- No subir a GitHub documentos de clientes, contraseñas, tokens, archivos `.env`, resultados ni el entorno virtual `.venv`.

## Publicar el código en GitHub
Crea un repositorio público vacío en https://github.com/new, con nombre `estandarizador-documental-tsf`.

Abre esta carpeta en VS Code y ejecuta:

```powershell
git init
git add .
git status
git commit -m "Publicar Estandarizador Documental TSF"
git branch -M main
$repo = Read-Host "Pega la URL HTTPS del repositorio de GitHub"
git remote add origin $repo
git push -u origin main
```

Si Git solicita una cuenta, completa el inicio de sesión de GitHub desde el navegador.

**Nota:** para subir código no es necesario habilitar GitHub Pages.
