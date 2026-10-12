# Portal de liquidación de comisiones · Grupo MG

Web estática (GitHub Pages) donde cada asesor, jefe de ventas y gerencia ve y descarga su liquidación.
No hay servidor: cada persona tiene un archivo cifrado en `d/` que solo se abre con su código personal.

## Qué ve cada perfil
- **Asesor**: solo su liquidación (monto, estado, ventas que la componen, evolución). Descarga Excel / imprime.
- **Jefe de ventas**: su comisión + la de su equipo (con detalle desplegable). Descarga Excel.
- **Gerencia comercial / general**: toda la empresa, filtro por equipo y búsqueda.

## Publicar (primera vez)
1. Crear un repositorio en la cuenta de GitHub del área (p. ej. `liquidacion-comisiones`).
2. Subir el contenido de esta carpeta (`index.html`, `lib/`, `d/`, `build_portal.py`, `tools/`, `README.md`, `.gitignore`).
3. Settings → Pages → Branch `main` / root. La URL queda `https://<cuenta>.github.io/<repo>/`.
4. Enviar a cada persona su código (hoja `CODIGOS_DE_ACCESO.xlsx`, **no se sube al repositorio**).

## Actualizar cada mes
```
pip install openpyxl cryptography msoffcrypto-tool
python3 build_portal.py "RESUMEN COMISIONES.xlsx" --cuadro "1 CUADRO DE VENTAS - COMISIONES.xlsx" --clave <contraseña del cuadro>
```
Luego subir la carpeta `d/` actualizada. Los códigos ya entregados se conservan (se leen de `CODIGOS_DE_ACCESO.xlsx`);
las personas nuevas reciben un código nuevo. Revisar `REPORTE_CONCILIACION.xlsx`: lista los casos donde el monto del
Resumen no coincide con el Cuadro de Ventas (en esos casos el portal muestra el monto del Resumen sin el detalle).

## Reglas
- El monto oficial es el del **Resumen de Comisiones**; el detalle de ventas viene del **Cuadro de Ventas** y solo se muestra si cuadra con el Resumen.
- Jefes: si el Cuadro tiene su bloque del mes se usa su total; si no, se muestra un monto *referencial* (unidades del equipo × Escala de Jefes), marcado "ref.".
- Corredores/empresas externas no tienen acceso propio, pero aparecen en la vista de Gerencia (lista `EXTERNOS` en `build_portal.py`).
- Equipos: se toman de los bloques SUM del Resumen; si no existen, de la tabla vigente `EQUIPOS` en `build_portal.py`.
- Cuentas de Gerencia: `GERENTE COMERCIAL` y `GERENTE GENERAL` (editar nombres en `GERENCIA`).
