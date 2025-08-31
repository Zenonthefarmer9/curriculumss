# mi_cv_proyecto

Plantilla de CV en .docx con python-docx (ATS-friendly + acento de color) y foto opcional.

## Estructura

- cv_template/
  - generate_cv.py: Script que genera el CV usando datos de ejemplo (demo_data).
  - batch_generate_cv.py: Generación por lotes desde JSON o Excel (con opción de procesar fotos PNG/JPG a 1:1, 600×600, <200 KB).
  - data/sample_profile.json: Ejemplo de perfil (opcional, no usado por el script actual).
  - data/profiles_sample.json: Ejemplo con múltiples perfiles.
  - assets/photos/: Coloca aquí tus fotos (README con recomendaciones).
- output/: Carpeta donde se guardan los .docx generados.
- requirements.txt: Dependencias de Python.

## Requisitos

- Python 3.9+
- Paquetes:
  - python-docx>=1.1.0
  - Opcionales (instala manualmente si los necesitas):
    - Pillow>=10.0.0 (procesado de fotos: recorte 1:1, 600×600, compresión <200KB)
    - pandas>=2.0.0 y openpyxl>=3.1.0 (lectura de Excel)

## Instalación de dependencias

```
cd mi_cv_proyecto
python -m pip install -r requirements.txt
# Para Excel:
python -m pip install pandas openpyxl
# Para procesar fotos:
python -m pip install Pillow
```

## Uso rápido (1 CV)

- Generar un CV de ejemplo (usa demo_data interna):
```
python cv_template/generate_cv.py
```
Salida: `output/CV_Natalia_Moreno_<Año>.docx`

- Personalizar datos:
  - Edita la función `demo_data()` dentro de `cv_template/generate_cv.py`.
  - Opcionalmente, cambia `incluir_foto` a `True` y ajusta `ruta_foto` a tu PNG/JPG.
  - Posición de foto: `photo_position` admite `right_paragraph` (por defecto), `right_table` o `left_table`.
  - LinkedIn: se filtra automáticamente de la línea de contacto.

## Generación por lotes (JSON)

- JSON de entrada: `cv_template/data/profiles_sample.json` (o uno propio con la misma estructura). Puedes usar PNG o JPG en `ruta_foto` y definir `incluir_foto: true`.

- Sin procesar fotos (inserta la imagen tal cual):
```
python cv_template/batch_generate_cv.py -i cv_template/data/profiles_sample.json -o output
```

- Procesando fotos para cumplir la plantilla (1:1, 600×600 px, <200 KB) — Requiere Pillow instalado manualmente:
```
python cv_template/batch_generate_cv.py \
  -i cv_template/data/profiles_sample.json \
  -o output \
  --process-photos
```
Esto generará copias normalizadas en `output/_photos_processed/` y usará esas versiones al crear los .docx.

## Generación por lotes (Excel)

- Prepara un Excel (primera hoja) con columnas sugeridas (no sensibles a mayúsculas):
  - nombre, cargo, email/correo, movil/celular/telefono, web, linkedin (opcional, se omitirá en la salida), ubicacion
  - foto/foto_filename/ruta_foto (opcional). Si no pones, se intenta buscar por el nombre del candidato en `assets/photos/`.
  - resumen
  - Para una experiencia simple: puesto, empresa, periodo/fecha, ubicacion_experiencia/ubicacion, sector, logros, actividades, proyectos.
  - O bien `experiencias_json` con una lista JSON de experiencias.
  - Para educación simple: grado, institucion/universidad, detalle (opcional).
  - O bien `educacion_json` con una lista JSON de items.
  - certificaciones, habilidades (listas separadas por coma o punto y coma)
  - idiomas (formato "Español:Nativo; Inglés:B2" o un JSON {"Español":"Nativo", ...})
  - photo_position (opcional): `right_paragraph`, `right_table`, `left_table`.

- Ejecuta:
```
python cv_template/batch_generate_cv.py -i cv_template/data/perfiles.xlsx -o output
```

- Con procesado de fotos (requiere Pillow):
```
python cv_template/batch_generate_cv.py -i cv_template/data/perfiles.xlsx -o output --process-photos
```

- Búsqueda automática de fotos:
  - Si `foto`/`ruta_foto` no está en Excel/JSON, se intentará localizar una foto en `cv_template/assets/photos/` cuyo nombre de archivo coincida (de forma flexible) con el nombre del candidato.
  - Se aceptan extensiones: .png, .jpg, .jpeg, .webp. El procesado convierte a .jpg para normalizar.

## Notas sobre la foto (opcional)

- Relación de aspecto: 1:1 (cuadrada).
- Resolución recomendada: 600×600 px (mínimo 400×400 px).
- Tamaño en el documento: ancho ~3.5 cm (alto proporcional).
- Fondo neutro, buena iluminación, encuadre hombros-cabeza.
- Peso ideal < 200 KB para mantener el .docx liviano.
- Nota ATS: si el proceso es 100% ATS estrictos, evita la foto; úsala solo para versiones “human-readable”.

## Troubleshooting

- Si ves `error: the following arguments are required: --input/-i`, ahora el script tiene un valor por defecto (`profiles_sample.json`). Ejecuta con `-i` si quieres otro archivo.
- Si usas `--process-photos` y no ves `output/_photos_processed/`, instala Pillow y vuelve a ejecutar.
- Para Excel, instala manualmente `pandas` y `openpyxl`.
- Si aparece `":" expected, got '1'` al leer un JSON, revisa que las claves estén entre comillas y no haya comas sobrantes al final (ya se corrigió `data/numeros_disponibles.json`).
