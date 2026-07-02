# Tablero Flores El Trigal — Ejecución Semanal (PWA)

Tablero de ejecución semanal de producción, ahora empaquetado como **Progressive Web App (PWA)**: se puede instalar en celular/escritorio con ícono propio y sigue funcionando (parcialmente) sin internet una vez visitada al menos una vez.

## Estructura de archivos

```
├── index.html        → estructura de la página (sin CSS/JS inline)
├── style.css          → todos los estilos
├── app.js             → toda la lógica: parsing del Excel, render, gráficos, exportación
├── manifest.json       → metadata de instalación (nombre, íconos, colores)
├── sw.js               → service worker: cachea el shell de la app para uso offline
├── icons/              → íconos en distintos tamaños (72 a 512px) + favicon + apple-touch-icon
└── README.md
```

El Excel del usuario **nunca se sube a ningún servidor**: se procesa 100% en el navegador de quien lo abre, con SheetJS.

## Publicar en GitHub Pages

1. Crea un repositorio nuevo en GitHub.
2. Sube **todos los archivos y la carpeta `icons/` completa**, manteniendo la misma estructura (no metas los archivos dentro de una subcarpeta).
   - Con "Add file" → "Upload files" puedes arrastrar todo de una vez, incluyendo la carpeta `icons`.
3. Ve a **Settings → Pages**.
4. **Source: Deploy from a branch** → Branch: `main` → carpeta `/ (root)` → **Save**.
5. En 1-2 minutos tendrás tu link: `https://tuusuario.github.io/nombre-repo/`.

## Qué gana con el formato PWA

- **Instalable**: en Chrome/Edge/Android aparece un botón "Instalar app" en la barra de direcciones; en iOS Safari se instala con "Compartir → Agregar a pantalla de inicio". Queda con ícono propio y abre en modo standalone (sin barra del navegador).
- **Funciona offline (parcial)**: el `sw.js` cachea el HTML/CSS/JS propios y las librerías de CDN (Bootstrap, SheetJS) la primera vez que se visita con internet. Después, aunque no haya señal, la app carga y se puede subir un Excel igual — solo necesita internet para la primera visita y para actualizaciones.
- **Más liviano de mantener**: CSS y JS separados del HTML, más fácil de editar o revisar diffs en GitHub.

## Actualizar el tablero más adelante

Si cambias `app.js` o `style.css` y vuelves a subir los archivos al repo:

- El **service worker detecta el cambio** y descarga la nueva versión en segundo plano, pero **el usuario puede seguir viendo la versión vieja cacheada hasta que recargue la página dos veces** (comportamiento normal de cualquier PWA con esta estrategia de caché).
- Si quieres forzar que todos noten la actualización de inmediato, sube el cambio y además incrementa `CACHE_VERSION` en `sw.js` (por ejemplo de `'v1'` a `'v2'`) — eso invalida el caché viejo automáticamente.

## Notas

- Requiere HTTPS para que el service worker funcione — GitHub Pages ya sirve todo por HTTPS por defecto, así que no hay que hacer nada extra.
- Si prefieres la versión simple de un solo archivo (sin instalación ni offline), esa sigue siendo válida — solo usa el `index.html` autocontenido de la versión anterior.
