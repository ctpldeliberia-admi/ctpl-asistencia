# CTPL Asistencia

Sistema de carnets y control de asistencia del personal del **Colegio Técnico Profesional de Liberia**.

- Registro unificado: carnet (datos y foto), cuenta y verificación de correo.
- Marca de entrada y salida con **QR dinámico** (cambia cada 15 s) y teléfono verificado.
- Horarios por día, días sin labores y **doble jornada** (diurna y nocturna).
- Informes semanales en app, correo, **PDF** y CSV.
- Panel de administración con impresión de carnets en hojas A4.

## Estructura

| Carpeta | Contenido |
|---|---|
| raíz (`index.html`, `admin.html`, `pantalla.html`, `css/`, `js/`, `img/`) | Aplicación web publicada con GitHub Pages |
| `apps-script/` | Servidor de datos (Google Apps Script + Google Sheets): `Code.gs`, `Asistencia.gs`, `Api.gs` |
| `manual/` | Manual de uso en PDF |

## Instalación

### 1. Servidor (cuenta Google del CTPL)
1. Cree una hoja de cálculo **CTPL Carnets y Asistencia** y abra **Extensiones › Apps Script**.
2. Cree tres archivos de secuencia de comandos: **Code**, **Asistencia** y **Api**, con el contenido de `apps-script/`.
3. Ejecute `setup` y luego `setupAsistencia` (autorice los permisos).
4. **Implementar › Nueva implementación › Aplicación web**: *Ejecutar como*: Yo · *Acceso*: Cualquier usuario.
5. Copie la URL que termina en `/exec`.

### 2. Sitio (GitHub Pages)
1. Pegue la URL del paso anterior en `config.js`.
2. **Settings › Pages › Deploy from a branch › main / (root) › Save**.
3. El sitio queda en `https://CUENTA.github.io/ctpl-asistencia/`.

### 3. Primer ingreso
1. Abra `admin.html` con la clave inicial y **cámbiela de inmediato** en **Configuración** (este repositorio es público).
2. En **Configuración** indique: dirección de la app, correos de la Dirección, jornada y ubicación del colegio.
3. Autorice la pantalla de marcas en `pantalla.html`.

## Actualizaciones
- **Servidor:** reemplace los archivos `.gs` y publique **Gestionar implementaciones › Nueva versión** (la URL no cambia).
- **Sitio:** suba los archivos modificados al repositorio; GitHub Pages publica automáticamente en 1 o 2 minutos.

## Enlaces
| Uso | Ruta |
|---|---|
| App del personal | `/` |
| Pantalla de marcas | `/pantalla.html` |
| Administración | `/admin.html` |

---
Uso interno · CTPL · Ministerio de Educación Pública
