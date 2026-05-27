# Referencias APA 7 — Complemento de Word (Office 365 web)

Complemento de Office.js que genera referencias y citas en formato **APA séptima edición** y las inserta directamente en tu documento de Word con sangría francesa y cursivas correctas.

![Vista del complemento](assets/icon-128.png)

## Funcionalidades

- 8 tipos de fuente: libro, capítulo, artículo de revista, periódico, sitio web, tesis, informe institucional y video.
- Genera la **referencia formateada** con cursivas correctas e inserta el párrafo con sangría francesa de 1.27 cm.
- Genera la **cita en texto** parentética y narrativa.
- Maneja correctamente las reglas APA 7: `et al.` con 3+ autores, listado completo hasta 20, elipsis con 21+.
- **Lista acumulada** de referencias guardadas en el navegador; inserta la bibliografía completa al final del documento ordenada alfabéticamente.
- Soporta autores institucionales con la sintaxis `[Nombre de la institución]`.

---

## Cómo instalarlo (paso a paso)

El complemento es 100% del lado del cliente, pero **Office 365 requiere que los archivos estén alojados en HTTPS**. La forma más fácil y gratuita es GitHub Pages.

### Paso 1 — Sube el proyecto a GitHub

1. Crea una cuenta en [github.com](https://github.com) si no tienes una.
2. Crea un nuevo repositorio **público** llamado, por ejemplo, `apa7-addin`.
3. Sube todos estos archivos manteniendo la estructura de carpetas:
   ```
   apa7-addin/
   ├── manifest.xml
   ├── README.md
   ├── assets/
   │   ├── icon-16.png
   │   ├── icon-32.png
   │   ├── icon-64.png
   │   ├── icon-80.png
   │   └── icon-128.png
   └── src/
       ├── taskpane.html
       ├── taskpane.css
       └── taskpane.js
   ```

### Paso 2 — Activa GitHub Pages

1. En el repositorio, abre **Settings → Pages**.
2. En **Source** elige la rama `main` (o `master`) y la carpeta `/ (root)`.
3. Guarda. Espera 1-2 minutos a que GitHub te dé una URL del tipo:
   ```
   https://TU_USUARIO.github.io/apa7-addin/
   ```
4. Verifica que abre: `https://TU_USUARIO.github.io/apa7-addin/src/taskpane.html` (deberías ver el panel en una página web).

### Paso 3 — Edita el manifest.xml

Abre `manifest.xml` y reemplaza **todas** las apariciones de `USUARIO.github.io` con tu URL real. Por ejemplo, si tu usuario es `mariolopez`, reemplaza `https://USUARIO.github.io/apa7-addin/` por `https://mariolopez.github.io/apa7-addin/`.

También genera un **GUID único** propio en [guidgenerator.com](https://www.guidgenerator.com/) y pégalo en `<Id>...</Id>`. Esto evita conflictos si compartes el complemento.

Sube el manifest editado a GitHub.

### Paso 4 — Carga el complemento en Word

1. Abre Word en [office.com](https://www.office.com) (versión navegador).
2. Abre cualquier documento.
3. Ve a la pestaña **Inicio → Complementos → Más complementos**.
4. En la ventana que se abre, pestaña **Mis complementos**, da clic en **Cargar mi complemento** (ubicado normalmente en la esquina superior derecha).
5. Selecciona el archivo `manifest.xml` desde tu computadora.
6. Aparecerá el botón **Referencias APA** en la cinta. Da clic para abrir el panel.

> **Para uso institucional o reutilizable:** un administrador de Microsoft 365 puede subir el manifest al Centro de administración → **Configuración → Aplicaciones integradas** para distribuirlo a toda la organización.

---

## Cómo usarlo

### Crear una referencia

1. **Pestaña "Crear"** → elige el tipo de fuente.
2. **Campo "Autores"**: escribe un autor por línea con el formato `Apellido, Nombre`.
   - Ejemplo: `García Márquez, Gabriel José`
   - Para una institución usa corchetes: `[Organización Mundial de la Salud]`
3. Llena los demás campos según el tipo de fuente.
4. Botones:
   - **Vista previa** muestra la referencia y la cita antes de insertar.
   - **Insertar referencia** pone el párrafo formateado donde está el cursor.
   - **Insertar cita** pega `(Apellido et al., 2023)` en el cursor.
   - **Guardar en lista** acumula la referencia para insertar la bibliografía completa después.

### Insertar la bibliografía completa

1. Guarda todas tus referencias con **Guardar en lista** a lo largo del documento.
2. Cuando termines, ve a la pestaña **Mis refs** y pulsa **Insertar lista de referencias ordenada**. Se inserta al final del documento un encabezado "Referencias" y la lista alfabética con sangría francesa.

### Reglas APA 7 que aplica automáticamente

| Autores | En referencia | En cita parentética | En cita narrativa |
|---|---|---|---|
| 1 | `Apellido, A. A.` | `(Apellido, 2023)` | `Apellido (2023)` |
| 2 | `Apellido, A. A., y Apellido, B. B.` | `(Apellido y Apellido, 2023)` | `Apellido y Apellido (2023)` |
| 3-20 | Todos listados con coma, "y" antes del último | `(Apellido et al., 2023)` | `Apellido et al. (2023)` |
| 21+ | Primeros 19, elipsis (`...`), último autor | `(Apellido et al., 2023)` | `Apellido et al. (2023)` |

---

## Estructura del proyecto

```
apa7-addin/
├── manifest.xml          ← Manifiesto del complemento (editar URLs y GUID)
├── README.md             ← Este archivo
├── assets/               ← Íconos PNG (16, 32, 64, 80, 128 px)
└── src/
    ├── taskpane.html     ← UI del panel
    ├── taskpane.css      ← Estilos (Fraunces + Inter)
    └── taskpane.js       ← Lógica APA 7 + Office.js
```

## Personalizar

- **Cambiar el color de acento**: edita la variable `--accent` en `src/taskpane.css`.
- **Agregar un nuevo tipo de fuente**: añade una entrada a `FIELDS_BY_TYPE` y una función `buildXxxxx()` en `src/taskpane.js`, y un `<option>` en el `<select id="ref-type">` de `taskpane.html`.
- **Cambiar "y" por "&"** en las referencias (algunas guías en español lo prefieren): busca `' y '` y `', y '` en `taskpane.js` y reemplaza por `' & '` y `', & '`.

## Limitaciones conocidas

- El complemento corre 100% en el cliente, no consulta bases bibliográficas externas (no busca DOIs en Crossref ni importa de Zotero). Es un formateador, no un gestor bibliográfico.
- Las referencias guardadas viven en el `localStorage` del navegador; si limpias datos del navegador o cambias de equipo, se pierden.
- Cumple las reglas más comunes de APA 7, pero los casos límite (autores con un solo apellido, obras antiguas reeditadas, podcasts, redes sociales, etc.) podrían requerir ajuste manual después de insertar.

## Licencia

Uso libre. Adáptalo a tus necesidades.
