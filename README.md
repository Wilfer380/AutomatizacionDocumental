# AutomatizaciónDocumental

Aplicación de escritorio en español para generar un DOCX y un PDF por serie desde:

- una plantilla Word que contenga `[Serie]`
- un libro de Excel con la lista de series

## Flujo de uso

1. Cargue la plantilla Word, el archivo Excel y la carpeta de salida.
2. Elija la hoja y la columna de serie.
3. Pulse **Validar** para precargar la vista previa.
4. Ajuste filtros, búsqueda y selección de filas.
5. Pulse **Generar documentos**.

## Qué incluye la interfaz

- Encabezado azul oscuro con título y ayuda
- Paneles numerados por sección
- Vista previa con resumen, búsqueda y tabla de filas
- Opciones de generación
- Barra de progreso con estado
- Acciones para generar, cancelar y abrir la carpeta de salida

## Requisitos

- Python 3.11+
- Windows recomendado para `docx2pdf`
- LibreOffice si `docx2pdf` no está disponible

## Instalación

```bash
pip install -r requirements.txt
```

## Ejecución

```bash
python run_app.py
```

## Notas

- La plantilla debe contener `[Serie]` antes de generar.
- La plantilla original no se modifica.
- Los archivos generados se organizan en `Word_generados`, `PDF_generados` y `Reportes` dentro de la carpeta elegida.
- Las rutas iniciales de ejemplo están pensadas para este equipo; ajústelas si trabajas en otra máquina.
