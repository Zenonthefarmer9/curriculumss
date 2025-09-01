# -*- coding: utf-8 -*-
"""
Generación por lotes de CVs en .docx usando exclusivamente cv_template/data/profiles.json como fuente de perfiles.
- Opcional: procesar/normalizar fotos (recorte 1:1, 600x600 px, compresión <200 KB) usando Pillow.

Uso ejemplos:
    python cv_template/batch_generate_cv.py -o output --process-photos

Notas:
    - Este script ignora cualquier --input y siempre usa data/profiles.json.
    - Usa merge_all_profiles.py si necesitas combinar varios JSON en profiles.json.
    - Si existe data/perfiles_generados_extra.json, se combinará en memoria (sin sobrescribir profiles.json).
"""

import argparse
import json
import os
import sys
import unicodedata
from typing import Any, Dict, List, Optional, Tuple

from generate_cv import construir_cv

PROJECT_ROOT = os.path.dirname(os.path.dirname(__file__))
PROCESSED_DIRNAME = os.path.join(PROJECT_ROOT, 'output', '_photos_processed')
DEFAULT_PHOTOS_DIR = os.path.join(PROJECT_ROOT, 'cv_template', 'assets', 'photos')
DEFAULT_PROFILES_FILE = os.path.join(PROJECT_ROOT, 'cv_template', 'data', 'profiles.json')
DEFAULT_EXTRA_FILE = os.path.join(PROJECT_ROOT, 'cv_template', 'data', 'perfiles_generados_extra.json')
DEFAULT_INPUT = DEFAULT_PROFILES_FILE


def load_json(path: str) -> Any:
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def normalize_profiles(data: Any) -> List[Dict[str, Any]]:
    if isinstance(data, dict) and 'perfiles' in data and isinstance(data['perfiles'], list):
        return data['perfiles']
    if isinstance(data, list):
        return data
    raise ValueError('El JSON debe ser una lista de perfiles o un objeto con clave "perfiles" (lista).')


REQUIRED_KEYS = ['nombre','cargo', 'contacto', 'resumen', 'experiencias', 'educacion', 'habilidades', 'idiomas']


def validate_profile(p: Dict[str, Any]) -> Tuple[bool, List[str]]:
    missing = [k for k in REQUIRED_KEYS if k not in p or p[k] in (None, '', [])]
    return (len(missing) == 0, missing)


def resolve_photo_path(ruta: Optional[str]) -> Optional[str]:
    if not ruta:
        return None
    if os.path.isabs(ruta):
        return ruta if os.path.exists(ruta) else None
    # relativa al root del proyecto
    candidate = os.path.join(PROJECT_ROOT, ruta)
    if os.path.exists(candidate):
        return candidate
    # relativa a assets/photos
    candidate2 = os.path.join(DEFAULT_PHOTOS_DIR, ruta)
    return candidate2 if os.path.exists(candidate2) else None


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def try_import_pillow():
    try:
        from PIL import Image  # noqa: F401
        return True
    except Exception:
        return False


def preprocess_photo(src_path: str, target_dir: str, target_size_px: int = 600, max_bytes: int = 200 * 1024) -> str:
    """
    - Convierte a RGB
    - Recorta centrado a cuadrado 1:1
    - Redimensiona a target_size_px x target_size_px
    - Comprime JPEG buscando <= max_bytes (calidad binaria 95..40)
    - Retorna la ruta del archivo procesado (JPG)
    """
    if not try_import_pillow():
        raise RuntimeError(
            'Pillow no está instalado. Instálalo manualmente (pip install Pillow) o ejecuta sin --process-photos.'
        )
    from PIL import Image

    ensure_dir(target_dir)

    base = os.path.splitext(os.path.basename(src_path))[0]
    out_path = os.path.join(target_dir, f"{base}_{target_size_px}.jpg")

    # Determinar filtro de remuestreo compatible con Pillow>=10 y anteriores
    try:
        resample = Image.Resampling.LANCZOS  # Pillow >= 10
    except Exception:
        resample = getattr(Image, 'LANCZOS', getattr(Image, 'ANTIALIAS', 1))

    with Image.open(src_path) as im:
        im = im.convert('RGB')
        w, h = im.size
        # recorte centrado a cuadrado
        if w != h:
            side = min(w, h)
            left = (w - side) // 2
            top = (h - side) // 2
            im = im.crop((left, top, left + side, top + side))
        # redimensionar
        im = im.resize((target_size_px, target_size_px), resample=resample)

        # compresión: búsqueda binaria de calidad
        low, high = 40, 95
        best_quality = high
        while low <= high:
            q = (low + high) // 2
            im.save(out_path, format='JPEG', quality=q, optimize=True)
            size = os.path.getsize(out_path)
            if size <= max_bytes:
                best_quality = q
                low = q + 1
            else:
                high = q - 1
        im.save(out_path, format='JPEG', quality=best_quality, optimize=True)

    return out_path


# ------------------
# Carga desde Excel
# ------------------

def _import_pandas_openpyxl():
    try:
        import pandas as pd  # noqa: F401
        import openpyxl  # noqa: F401
        return True
    except Exception:
        return False


def _to_bool(val: Any) -> bool:
    if isinstance(val, bool):
        return val
    if val is None:
        return False
    s = str(val).strip().lower()
    return s in ('1', 'true', 'sí', 'si', 'y', 'yes', 'x')


def _split_list(val: Any) -> List[str]:
    if val is None:
        return []
    s = str(val).strip()
    if not s:
        return []
    # admitir separadores ; o ,
    parts = [p.strip() for p in s.replace('\n', ';').replace(',', ';').split(';')]
    return [p for p in parts if p]


def _parse_idiomas(val: Any) -> Dict[str, str]:
    # Formatos aceptados: "Español:Nativo;Inglés:B2" o JSON {"Español":"Nativo",...}
    if val is None:
        return {}
    s = str(val).strip()
    if not s:
        return {}
    try:
        obj = json.loads(s)
        if isinstance(obj, dict):
            return {str(k): str(v) for k, v in obj.items()}
    except Exception:
        pass
    result: Dict[str, str] = {}
    for item in _split_list(s):
        if ':' in item:
            k, v = item.split(':', 1)
            result[k.strip()] = v.strip()
        else:
            result[item] = ''
    return result


def _parse_experiencias(row: Dict[str, Any]) -> List[Dict[str, Any]]:
    # Opción 1: columna experiencias_json
    exp_json = row.get('experiencias_json') or row.get('experiencias')
    if exp_json:
        try:
            data = json.loads(exp_json) if isinstance(exp_json, str) else exp_json
            if isinstance(data, list):
                return data
        except Exception:
            pass
    # Opción 2: columnas planas para una sola experiencia
    puesto = row.get('puesto') or row.get('cargo_experiencia')
    empresa = row.get('empresa')
    periodo = row.get('periodo') or row.get('fecha') or ''
    ubicacion = row.get('ubicacion_experiencia') or row.get('ubicacion')
    sector = row.get('sector')
    if not (puesto and empresa):
        return []
    logros = _split_list(row.get('logros'))
    actividades = _split_list(row.get('actividades'))
    proyectos = _split_list(row.get('proyectos'))
    return [{
        'puesto': puesto,
        'empresa': empresa,
        'periodo': periodo,
        'ubicacion': ubicacion,
        'sector': sector,
        'logros': logros,
        'actividades': actividades,
        'proyectos': proyectos,
    }]


def _parse_educacion(row: Dict[str, Any]) -> List[Dict[str, Any]]:
    ed_json = row.get('educacion_json') or row.get('educacion')
    if ed_json:
        try:
            data = json.loads(ed_json) if isinstance(ed_json, str) else ed_json
            if isinstance(data, list):
                return data
        except Exception:
            pass
    # columnas simples
    grado = row.get('grado')
    inst = row.get('institucion') or row.get('universidad')
    detalle = row.get('detalle')
    if not (grado and inst):
        return []
    return [{'grado': grado, 'institucion': inst, 'detalle': detalle}]


def _slug(s: str) -> str:
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = s.lower().strip()
    # quitar caracteres no alfanuméricos
    out = []
    for ch in s:
        if ch.isalnum():
            out.append(ch)
        elif ch in (' ', '-', '_'):
            out.append('-')
    # colapsar guiones
    slug = ''.join(out)
    while '--' in slug:
        slug = slug.replace('--', '-')
    return slug.strip('-')


def _find_photo_by_basename(photos_dir: str, base: str) -> Optional[str]:
    if not base:
        return None
    stem, ext = os.path.splitext(base)
    candidates: List[str] = []
    if ext:
        p = os.path.join(photos_dir, base)
        if os.path.exists(p):
            return p
    else:
        for e in ('.png', '.jpg', '.jpeg', '.webp'):
            p = os.path.join(photos_dir, stem + e)
            if os.path.exists(p):
                candidates.append(p)
        if candidates:
            # priorizar png/jpg
            for pref in ('.png', '.jpg', '.jpeg', '.webp'):
                for c in candidates:
                    if c.lower().endswith(pref):
                        return c
    return None


def _find_photo_by_name_guess(photos_dir: str, nombre: str) -> Optional[str]:
    if not nombre:
        return None
    slug = _slug(nombre)
    if not os.path.isdir(photos_dir):
        return None
    exts = ('.png', '.jpg', '.jpeg', '.webp')
    try:
        for fn in os.listdir(photos_dir):
            stem, ext = os.path.splitext(fn)
            if ext.lower() in exts:
                s_stem = _slug(stem)
                if slug in s_stem or s_stem in slug:
                    cand = os.path.join(photos_dir, fn)
                    if os.path.exists(cand):
                        return cand
    except Exception:
        return None
    return None


def load_profiles_from_excel(path: str, photos_dir: str = DEFAULT_PHOTOS_DIR) -> List[Dict[str, Any]]:
    if not _import_pandas_openpyxl():
        raise RuntimeError('Leer Excel requiere pandas y openpyxl. Instálalos manualmente: pip install pandas openpyxl')
    import pandas as pd

    # Leer primera hoja
    df = pd.read_excel(path)
    # normalizar nombres de columnas a minúsculas sin espacios
    df.columns = [str(c).strip().lower() for c in df.columns]

    profiles: List[Dict[str, Any]] = []

    for _, r in df.iterrows():
        row = r.to_dict()
        # nombre completo obligatorio
        nombre = row.get('nombre') or 'Sin Nombre'
        cargo = row.get('cargo') or ''

        # contacto: email, móvil, web, linkedin (linkedin será filtrado luego en generate_cv)
        email = row.get('email') or row.get('correo')
        movil = row.get('movil') or row.get('celular') or row.get('telefono')
        web = row.get('web') or row.get('sitio')
        linkedin = row.get('linkedin')
        contacto = [v for v in [email, movil, linkedin, web] if v]

        ubicacion = row.get('ubicacion')

        # foto
        incluir_foto = _to_bool(row.get('incluir_foto'))
        foto_filename = row.get('foto') or row.get('foto_filename') or row.get('ruta_foto')
        ruta_foto = None
        if foto_filename:
            ruta_foto = _find_photo_by_basename(photos_dir, str(foto_filename).strip())
            if ruta_foto:
                ruta_foto = os.path.relpath(ruta_foto, PROJECT_ROOT)
                incluir_foto = True
        if not ruta_foto:
            guess = _find_photo_by_name_guess(photos_dir, nombre)
            if guess:
                ruta_foto = os.path.relpath(guess, PROJECT_ROOT)
                incluir_foto = True

        resumen = row.get('resumen') or ''
        experiencias = _parse_experiencias(row)
        educacion = _parse_educacion(row)
        certificaciones = _split_list(row.get('certificaciones'))
        habilidades = _split_list(row.get('habilidades'))
        idiomas = _parse_idiomas(row.get('idiomas'))

        profile = {
            'nombre': nombre,
            'cargo': cargo,
            'contacto': contacto,
            'ubicacion': ubicacion,
            'incluir_foto': incluir_foto,
            'ruta_foto': ruta_foto,
            'resumen': resumen,
            'experiencias': experiencias,
            'educacion': educacion,
            'certificaciones': certificaciones,
            'habilidades': habilidades,
            'idiomas': idiomas,
            # permitir configurar posición de foto si existe columna
            'photo_position': row.get('photo_position') or 'right_paragraph',
        }
        profiles.append(profile)

    return profiles


def build_one(profile: Dict[str, Any], outdir: str, process_photos: bool, target_px: int, max_bytes: int):
    ok, missing = validate_profile(profile)
    if not ok:
        print(f"[WARN] Perfil omitido, faltan campos requeridos: {missing}")
        return

    data = dict(profile)  # copia superficial

    ruta_foto = resolve_photo_path(profile.get('ruta_foto'))
    incluir_foto = bool(profile.get('incluir_foto'))

    if incluir_foto and not ruta_foto:
        # intento por nombre
        guess = _find_photo_by_name_guess(DEFAULT_PHOTOS_DIR, profile.get('nombre', ''))
        if guess:
            ruta_foto = guess
        else:
            print("[INFO] 'incluir_foto' es True pero no se encontró 'ruta_foto' válida. Se omitirá la foto.")

    if incluir_foto and ruta_foto and process_photos:
        try:
            processed = preprocess_photo(ruta_foto, PROCESSED_DIRNAME, target_size_px=target_px, max_bytes=max_bytes)
            data['ruta_foto'] = processed
            print(f"[INFO] Foto procesada: {processed}")
        except Exception as e:
            print(f"[WARN] No se pudo procesar la foto '{ruta_foto}': {e}. Se intentará usar la original.")
            data['ruta_foto'] = ruta_foto
    else:
        data['ruta_foto'] = ruta_foto

    ensure_dir(outdir)
    construir_cv(data, carpeta_salida=outdir)


# ------------------
# Utilidad: unir JSON extra -> profiles.json
# ------------------

def _normalize_perfiles(data: Any) -> List[Dict[str, Any]]:
    if isinstance(data, dict) and isinstance(data.get('perfiles'), list):
        return data['perfiles']
    if isinstance(data, list):
        return data
    raise ValueError('El archivo debe ser lista o {"perfiles": [...]}')


def _profile_key_for_merge(p: Dict[str, Any]) -> Tuple[str, str, str]:
    nombre = str(p.get('nombre', '')).strip().lower()
    cargo = str(p.get('cargo', '')).strip().lower()
    try:
        fingerprint = json.dumps(p, ensure_ascii=False, sort_keys=True)
    except Exception:
        fingerprint = nombre + '|' + cargo
    return (nombre, cargo, fingerprint)


def merge_extra_into_profiles(extra_path: str, target_path: str) -> int:
    if not os.path.exists(extra_path):
        return 0
    with open(extra_path, 'r', encoding='utf-8') as f:
        extra_raw = json.load(f)
    with open(target_path, 'r', encoding='utf-8') as f:
        target_raw = json.load(f)
    extra = _normalize_perfiles(extra_raw)
    target = _normalize_perfiles(target_raw)
    combined = list(target) + list(extra)
    seen = set()
    result: List[Dict[str, Any]] = []
    for p in combined:
        k = _profile_key_for_merge(p)
        if k in seen:
            continue
        seen.add(k)
        result.append(p)
    with open(target_path, 'w', encoding='utf-8') as f:
        json.dump({'perfiles': result}, f, ensure_ascii=False, indent=2)
    return len(result)


def main():
    parser = argparse.ArgumentParser(description='Genera múltiples CVs .docx desde data/profiles.json')
    # --input queda oculto/ignorado para compatibilidad pero no se usará
    parser.add_argument('--input', '-i', default=DEFAULT_INPUT, help='(IGNORADO) Siempre se usa data/profiles.json')
    parser.add_argument('--outdir', '-o', default=os.path.join(PROJECT_ROOT, 'output'), help='Carpeta de salida (default: output)')
    parser.add_argument('--photos-dir', default=DEFAULT_PHOTOS_DIR, help='Carpeta donde buscar fotos por nombre (Excel)')
    parser.add_argument('--process-photos', action='store_true', help='Procesar fotos (recorte 1:1, 600x600, compresión <200KB) con Pillow')
    parser.add_argument('--target-size', type=int, default=600, help='Tamaño objetivo del lado de la imagen (px)')
    parser.add_argument('--max-bytes', type=int, default=200*1024, help='Tamaño máximo en bytes tras compresión (default 200KB)')
    # Merge opcional (en disco, previo a generar)
    parser.add_argument('--merge-extra-to-profiles', action='store_true', help='Unir data/perfiles_generados_extra.json dentro de data/profiles.json antes de generar')
    parser.add_argument('--extra-file', default=DEFAULT_EXTRA_FILE, help='Ruta del JSON extra (default: data/perfiles_generados_extra.json)')
    parser.add_argument('--profiles-file', default=DEFAULT_PROFILES_FILE, help='Ruta del profiles.json destino (default: data/profiles.json)')
    args = parser.parse_args()

    if args.merge_extra_to_profiles:
        try:
            total = merge_extra_into_profiles(args.extra_file, args.profiles_file)
            print(f"[INFO] Merge completado en '{args.profiles_file}'. Total perfiles: {total}")
        except Exception as e:
            print(f"[ERROR] Falló el merge de perfiles: {e}")
            sys.exit(1)

    # Forzar uso exclusivo de profiles.json
    input_path = DEFAULT_PROFILES_FILE
    print(f"[INFO] Usando única fuente: {os.path.relpath(input_path, PROJECT_ROOT)}")
    print(f"[INFO] Output dir: {args.outdir}")
    print(f"[INFO] Process photos: {args.process_photos}")

    if not os.path.exists(input_path):
        print(f"[ERROR] No existe {input_path}. Combina tus archivos con merge_all_profiles.py o crea profiles.json válido.")
        sys.exit(1)

    try:
        data = load_json(input_path)
        perfiles = normalize_profiles(data)
    except Exception as e:
        print(f"[ERROR] No se pudo leer perfiles: {e}")
        sys.exit(1)

    # Combinar en memoria con el archivo extra si existe (sin modificar profiles.json en disco)
    if os.path.exists(DEFAULT_EXTRA_FILE):
        try:
            extra_doc = load_json(DEFAULT_EXTRA_FILE)
            perfiles_extra = normalize_profiles(extra_doc)
            before = len(perfiles)
            combined = perfiles + perfiles_extra
            seen = set()
            dedup: List[Dict[str, Any]] = []
            for p in combined:
                k = _profile_key_for_merge(p)
                if k in seen:
                    continue
                seen.add(k)
                dedup.append(p)
            perfiles = dedup
            print(f"[INFO] Añadidos {len(perfiles) - before} perfiles extra desde {os.path.relpath(DEFAULT_EXTRA_FILE, PROJECT_ROOT)}. Total: {len(perfiles)}")
        except Exception as e:
            print(f"[WARN] No se pudo combinar el extra: {e}")

    if args.process_photos and not try_import_pillow():
        print('[ERROR] --process-photos requiere Pillow. Instálalo manualmente: pip install Pillow')
        sys.exit(2)

    count = 0
    for p in perfiles:
        try:
            build_one(p, args.outdir, args.process_photos, args.target_size, args.max_bytes)
            count += 1
        except Exception as e:
            print(f"[ERROR] Falló la generación de un CV: {e}")

    print(f"Listo. CVs generados: {count} -> {args.outdir}")


if __name__ == '__main__':
    main()
