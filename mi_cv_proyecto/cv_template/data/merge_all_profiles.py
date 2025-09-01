# -*- coding: utf-8 -*-
"""
Combina múltiples archivos JSON de perfiles en un único profiles.json.

Uso rápido (combinar automáticamente todos los JSON de perfiles en data/):
    python cv_template/data/merge_all_profiles.py --all

Uso con lista de archivos específica:
    python cv_template/data/merge_all_profiles.py --inputs cv_template/data/profiles_sample.json cv_template/data/perfiles_generados_extra.json

Opciones:
    --target: salida (por defecto: cv_template/data/profiles.json)

Reglas:
- Acepta formatos: {"perfiles": [...]}, lista de perfiles, o un único perfil como objeto.
- Deduplica por (nombre, cargo) con fallback a huella JSON.
- Ignora JSON que no contengan perfiles válidos (p. ej., catálogos auxiliares como numeros_disponibles.json).
"""
import argparse
import glob
import json
import os
import sys
from typing import Any, Dict, Iterable, List, Tuple

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
DATA_DIR = os.path.join(PROJECT_ROOT, 'cv_template', 'data')
DEFAULT_TARGET = os.path.join(DATA_DIR, 'profiles.json')

# Archivos típicos a excluir cuando se usa --all
EXCLUDE_BASENAMES = {
    'profiles.json',
    'profiles_sample.json',
    'sample_profile.json',
    'numeros_disponibles.json',
}


def _load_json(path: str) -> Any:
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def _normalize_perfiles(doc: Any) -> List[Dict[str, Any]]:
    # {"perfiles": [...]}
    if isinstance(doc, dict) and isinstance(doc.get('perfiles'), list):
        return [p for p in doc['perfiles'] if isinstance(p, dict)]
    # lista directa
    if isinstance(doc, list):
        return [p for p in doc if isinstance(p, dict)]
    # objeto único
    if isinstance(doc, dict) and ('nombre' in doc or 'cargo' in doc):
        return [doc]
    return []


def _profile_key(p: Dict[str, Any]) -> Tuple[str, str, str]:
    nombre = str(p.get('nombre', '')).strip().lower()
    cargo = str(p.get('cargo', '')).strip().lower()
    try:
        fingerprint = json.dumps(p, ensure_ascii=False, sort_keys=True)
    except Exception:
        fingerprint = nombre + '|' + cargo
    return (nombre, cargo, fingerprint)


def _iter_candidate_files() -> Iterable[str]:
    pattern = os.path.join(DATA_DIR, '*.json')
    for path in glob.glob(pattern):
        base = os.path.basename(path)
        if base in EXCLUDE_BASENAMES:
            continue
        yield path


def merge_files(input_files: List[str], target_path: str) -> int:
    all_profiles: List[Dict[str, Any]] = []

    for inp in input_files:
        try:
            doc = _load_json(inp)
            perf = _normalize_perfiles(doc)
            if not perf:
                print(f"[INFO] Sin perfiles válidos, se omite: {inp}")
                continue
            all_profiles.extend(perf)
            print(f"[OK] {len(perf)} perfiles desde {os.path.basename(inp)}")
        except Exception as e:
            print(f"[WARN] No se pudo leer {inp}: {e}")

    # si ya existe target, incluirlo primero para priorizar su orden
    if os.path.exists(target_path):
        try:
            doc = _load_json(target_path)
            perf = _normalize_perfiles(doc)
            if perf:
                all_profiles = perf + all_profiles
                print(f"[OK] Se preservan {len(perf)} perfiles existentes en {os.path.basename(target_path)}")
        except Exception as e:
            print(f"[WARN] No se pudo leer existente {target_path}: {e}")

    # deduplicar
    seen = set()
    result: List[Dict[str, Any]] = []
    for p in all_profiles:
        k = _profile_key(p)
        if k in seen:
            continue
        seen.add(k)
        result.append(p)

    os.makedirs(os.path.dirname(target_path), exist_ok=True)
    with open(target_path, 'w', encoding='utf-8') as f:
        json.dump({'perfiles': result}, f, ensure_ascii=False, indent=2)

    print(f"[DONE] Perfiles combinados: {len(result)} -> {target_path}")
    return len(result)


def main():
    parser = argparse.ArgumentParser(description='Combina múltiples JSON de perfiles en profiles.json')
    parser.add_argument('--inputs', nargs='*', help='Lista de archivos JSON de entrada')
    parser.add_argument('--all', action='store_true', help='Buscar automáticamente todos los JSON de perfiles en data/')
    parser.add_argument('--target', default=DEFAULT_TARGET, help='Salida profiles.json (default: data/profiles.json)')
    args = parser.parse_args()

    inputs: List[str] = []
    if args.all:
        inputs = list(_iter_candidate_files())
        if not inputs:
            print('[ERROR] No se encontraron archivos candidatos en data/.')
            sys.exit(1)
    elif args.inputs:
        for p in args.inputs:
            inputs.append(p if os.path.isabs(p) else os.path.join(PROJECT_ROOT, p))
    else:
        print('[ERROR] Especifica --all o --inputs ...')
        sys.exit(1)

    # normalizar target
    target = args.target if os.path.isabs(args.target) else os.path.join(PROJECT_ROOT, args.target)

    merge_files(inputs, target)


if __name__ == '__main__':
    main()
