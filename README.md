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

## Fase 2 - Gestión de Dossier

La aplicación ahora incluye una pestaña separada para preparar el flujo de dossier en modo simulación o en modo real confirmado.

Incluye:

- Configuración JSON local con ejemplo reutilizable y root documental editable
- Lectura flexible de columnas `CP` y `Serie` desde Excel
- Validación flexible de la carpeta `06_DOSSIER`, `Planos` y la Serie dentro de `Planos`
- Cuatro documentos PDF configurables desde la UI con reglas de carpeta `5/6/7`
- Copia real con confirmación explícita y respaldo previo antes de reemplazar
- Resultado JSON/log-friendly para auditoría y soporte

### Modo simulación

- No escribe archivos.
- Sólo planifica destinos, valida rutas y genera reporte.

### Modo real

- Requiere confirmación explícita del operador.
- Crea respaldo del destino antes de reemplazar.
- Escribe sólo dentro del CP validado.

### Cautelas operativas

- Revise el root documental antes de ejecutar.
- Revise la lista de PDFs fuente antes de confirmar.
- No use modo real sobre una carpeta compartida si la validación mostró ambigüedad.
- Si un CP, `06_DOSSIER`, `Planos` o la Serie no coinciden, la fila se omite y el lote continúa.

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
- `config.json` contiene la configuración local de Fase 2 y se mantiene fuera de control de versiones.
