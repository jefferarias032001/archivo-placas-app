# 🚛 Archivo de Placas — Web App

Consulta el parque automotor de la empresa: busca por **placa, conductor, tenedor o propietario** y filtra por **tipología** (tractomula, turbo, sencillo, patineta…), carrocería, marca y estado.

## Cómo funciona este proyecto

```
OneDrive (ARCHIVO DE PLACAS.xlsx)
        │
        │  actualizar.bat  ← UN SOLO COMANDO (doble clic)
        ▼
convertir.py lee el Excel → genera data.js (datos limpios y sin duplicados)
        ▼
git commit + git push → el repositorio en GitHub queda actualizado
        ▼
index.html abre y muestra los datos de inmediato (sin cargar nada)
```

| Archivo | Qué es |
|---|---|
| `index.html` | La aplicación completa. Doble clic y abre con los datos. |
| `data.js` | Los datos convertidos. Lo genera `convertir.py`. **No editar a mano.** |
| `convertir.py` | Lee el Excel de OneDrive, quita duplicados (deja el registro más reciente por placa) y genera `data.js`. |
| `actualizar.bat` | El comando único: convierte + commit + push. |
| `.gitignore` | Evita que el Excel original se suba por accidente. |

## 🔒 MUY IMPORTANTE — Privacidad

`data.js` contiene cédulas, celulares y direcciones de conductores, tenedores y propietarios. Por eso:

- El repositorio en GitHub **DEBE SER PRIVADO**. Nunca lo cambies a público.
- **No actives GitHub Pages** (en el plan gratis solo funciona con repos públicos, lo que expondría los datos de todas esas personas en internet). La app no lo necesita: funciona con doble clic en `index.html`.
- Comparte el acceso al repo solo con compañeros autorizados (en GitHub: Settings → Collaborators → Add people).

---

# Guía paso a paso desde cero

## Parte 1 — Instalar herramientas (solo una vez)

### 1.1 Git
1. Descarga de <https://git-scm.com/downloads> e instala **dejando todo por defecto**.
2. Comprueba: abre el menú inicio, escribe `cmd`, abre el "Símbolo del sistema" y escribe `git --version`. Debe responder algo como `git version 2.45.0`.

### 1.2 Python
`convertir.py` necesita Python para leer el Excel.
1. Descarga de <https://www.python.org/downloads> (botón amarillo).
2. Al instalar, **marca la casilla "Add Python to PATH"** (¡muy importante!) y luego "Install Now".
3. Comprueba en cmd: `python --version`.
4. Instala la librería que lee Excel. En el mismo cmd:
   ```
   pip install openpyxl
   ```

### 1.3 Visual Studio Code
Descarga e instala de <https://code.visualstudio.com>. Opcional: instala la extensión "Spanish Language Pack" para verlo en español.

### 1.4 Cuenta de GitHub
Crea una en <https://github.com> → Sign up.

### 1.5 Identifícate en Git (solo una vez)
En cmd, con tus datos reales:
```
git config --global user.name "Tu Nombre"
git config --global user.email "tucorreo@ejemplo.com"
```

## Parte 2 — Crear el repositorio PRIVADO en GitHub

1. Entra a <https://github.com/new>.
2. **Repository name:** `archivo-placas-app`
3. **Visibilidad: marca "Private"** ✅ (obligatorio — ver sección de privacidad).
4. No marques ninguna casilla de inicialización. Clic en **Create repository**.
5. Copia la dirección que aparece, tipo: `https://github.com/TU-USUARIO/archivo-placas-app.git`

## Parte 3 — Configurar el proyecto en tu PC

1. Descomprime el ZIP en una carpeta fija, por ejemplo `C:\Proyectos\archivo-placas-app`.

2. Abre **VS Code** → **Archivo → Abrir carpeta** → esa carpeta.

3. Abre `convertir.py` y edita la línea de configuración (está al inicio, bien señalada) con la ruta REAL de tu Excel en OneDrive:
   ```python
   RUTA_EXCEL = "C:/Users/TU_USUARIO/OneDrive - TU_EMPRESA/ARCHIVO DE PLACAS.xlsx"
   ```
   💡 Truco para sacar la ruta exacta: en el Explorador de Windows, mantén **Shift** y haz clic derecho sobre el archivo → "Copiar como ruta de acceso". Pégala y cambia las `\` por `/`.
   💡 Esto funciona porque OneDrive sincroniza el archivo en tu PC. Si tu OneDrive está "solo en línea", haz clic derecho sobre el archivo → "Conservar siempre en este dispositivo".

4. Abre la terminal de VS Code (**Terminal → Nueva terminal**) y ejecuta uno por uno:
   ```
   git init
   git add .
   git commit -m "Primera version de la web app de placas"
   git branch -M main
   git remote add origin https://github.com/TU-USUARIO/archivo-placas-app.git
   git push -u origin main
   ```
   (⚠️ en la quinta línea pon tu dirección real; la primera vez que hagas `push` se abrirá el navegador para iniciar sesión en GitHub)

5. Recarga la página del repo en GitHub: deben verse los archivos. 🎉

## Parte 4 — El comando único de actualización

Cada vez que el Excel cambie en OneDrive:

➡️ **Doble clic en `actualizar.bat`** (o en la terminal de VS Code: `.\actualizar.bat`)

Eso hace las tres cosas: convierte el Excel a `data.js`, hace el commit con la fecha y lo sube a GitHub. Al terminar verás "✅ Listo".

Luego abres `index.html` y ya está todo actualizado. En la barra superior aparece la fecha de la última actualización.

## Parte 5 — Uso diario de la app

- **Doble clic en `index.html`** → abre con los datos al instante.
- **Buscador:** placa, nombre, cédula o celular de conductor, tenedor o propietario.
- **Tipología:** menú agrupado (TRACTOMULA, TURBO, SENCILLO, PATINETA, DOBLETROQUE, ZORRO…) + botones rápidos con el conteo de cada una. La tarjeta muestra la tipología exacta (ej. `SENCILLO_ESTACAS`).
- **Carrocería, Marca, Estado** (activo/inactivo).
- **"Ver detalle"** en cada tarjeta: cédulas, direcciones, licencia del conductor y vencimiento.
- Los celulares son enlaces: desde el celular, un toque y llama.
- **"Exportar filtrado (CSV)"**: descarga lo que estés viendo para abrirlo en Excel (recuerda que contiene datos personales).

## Si otro compañero quiere usar la app

1. Que cree su cuenta de GitHub y tú lo agregues en: repo → Settings → Collaborators.
2. En su PC: instala Git, abre cmd en una carpeta y ejecuta:
   ```
   git clone https://github.com/TU-USUARIO/archivo-placas-app.git
   ```
3. Doble clic en `index.html`. Para traer datos nuevos cuando tú actualices: `git pull`.

## Comandos útiles

| Comando | Para qué sirve |
|---|---|
| `git status` | Ver qué archivos cambiaron |
| `git log --oneline` | Ver el historial de actualizaciones |
| `git pull` | Traer lo último que se subió a GitHub |
