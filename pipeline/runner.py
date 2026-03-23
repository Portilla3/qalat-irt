"""
pipeline/runner_irt.py — Runner para la app QALAT IRT
Usa subprocess + variables de entorno para todos los scripts.
"""
import sys, os, re, tempfile, shutil, types
from io import BytesIO
from pathlib import Path

PIPELINE_DIR = Path(__file__).parent

OUTPUTS = {
    'caract_excel': ('IRT_Caracterizacion_Ingreso.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
    'seg_excel':    ('IRT_Seguimiento.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
    'word_caract':  ('IRT_Informe_Caracterizacion.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
    'word_seg':     ('IRT_Informe_Seguimiento.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
}

SCRIPT_FILES = {
    'caract_excel': 'caract_excel.py',
    'seg_excel':    'seg_excel.py',
    'word_caract':  'word_caract.py',
    'word_seg':     'word_seg.py',
}


def _patch_and_run(script_key, wide_path, out_path, filtro_centro=None):
    """
    Parchea el script con env vars y lo ejecuta via subprocess.
    Todos los scripts IRT usan el mismo patrón.
    """
    import subprocess

    src = open(str(PIPELINE_DIR / SCRIPT_FILES[script_key]), encoding='utf-8').read()

    # Parchar INPUT_FILE
    src = re.sub(
        r'INPUT_FILE\s*=\s*auto_archivo_wide\(\)',
        'INPUT_FILE = __import__("os").environ["QALAT_WIDE"]',
        src
    )
    src = re.sub(
        r'INPUT_FILE\s*=\s*None\s*#.*runner.*',
        'INPUT_FILE = __import__("os").environ["QALAT_WIDE"]',
        src
    )
    # Parchar OUTPUT_FILE
    src = re.sub(
        r"OUTPUT_FILE\s*=\s*'/home/claude/[^']*\.(?:xlsx|docx)'",
        'OUTPUT_FILE = __import__("os").environ["QALAT_OUT"]',
        src
    )
    src = re.sub(
        r'OUTPUT_FILE\s*=\s*None\s*#.*runner.*',
        'OUTPUT_FILE = __import__("os").environ["QALAT_OUT"]',
        src
    )
    # Parchar FILTRO_CENTRO
    src = re.sub(
        r'FILTRO_CENTRO\s*=\s*None\s*#.*runner.*',
        'FILTRO_CENTRO = __import__("os").environ.get("QALAT_CENTRO") or None',
        src
    )
    src = re.sub(
        r'^FILTRO_CENTRO\s*=\s*None\s*$',
        'FILTRO_CENTRO = __import__("os").environ.get("QALAT_CENTRO") or None',
        src, flags=re.MULTILINE
    )

    fd, tmp_py = tempfile.mkstemp(suffix='.py', prefix='qalat_irt_')
    os.close(fd)
    with open(tmp_py, 'w', encoding='utf-8') as f:
        f.write(src)

    env = os.environ.copy()
    env['QALAT_WIDE']   = wide_path
    env['QALAT_OUT']    = out_path
    env['QALAT_CENTRO'] = filtro_centro or ''

    try:
        r = subprocess.run(
            [sys.executable, tmp_py],
            capture_output=True, text=True,
            timeout=180, env=env
        )
        if r.returncode != 0:
            raise RuntimeError(r.stderr[-2000:] or r.stdout[-2000:])
    finally:
        try: os.unlink(tmp_py)
        except: pass


def run_script(script_key, wide_path, filtro_centro=None):
    out_filename, mimetype = OUTPUTS[script_key]
    if filtro_centro:
        base, ext = out_filename.rsplit('.', 1)
        out_filename = f'{base}_{filtro_centro}.{ext}'

    suffix = '.' + out_filename.rsplit('.', 1)[1]
    fd, out_path = tempfile.mkstemp(suffix=suffix, prefix='qalat_irt_out_')
    os.close(fd)

    try:
        _patch_and_run(script_key, wide_path, out_path, filtro_centro)

        if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
            raise FileNotFoundError('El script no generó salida')

        with open(out_path, 'rb') as f:
            data = f.read()
        return BytesIO(data), out_filename, mimetype

    finally:
        try: os.unlink(out_path)
        except: pass


def run_all(wide_path, progress_cb=None):
    results = {}
    keys = list(OUTPUTS.keys())
    for i, key in enumerate(keys):
        if progress_cb: progress_cb(i, len(keys), key)
        try:
            buf, fname, mime = run_script(key, wide_path)
            results[key] = {'ok': True, 'buf': buf, 'fname': fname, 'mime': mime}
        except Exception as e:
            results[key] = {'ok': False, 'error': str(e)}
    if progress_cb: progress_cb(len(keys), len(keys), 'listo')
    return results


# ── Distribución por centros ──────────────────────────────────────────────────
import zipfile, unicodedata as _ud

def _slug(s):
    s = _ud.normalize('NFD', str(s)).encode('ascii', 'ignore').decode()
    s = re.sub(r'[^\w\s-]', '', s).strip()
    s = re.sub(r'[\s]+', '_', s)
    return s[:60]

def _detectar_centros(wide_path):
    import pandas as pd
    try:
        df = pd.read_excel(wide_path, sheet_name='Por Centro', header=2)
        col = df.columns[0]
        return [str(v).strip() for v in df[col].dropna()
                if str(v).strip().upper() != 'TOTAL' and str(v).strip()]
    except:
        df = pd.read_excel(wide_path, sheet_name='Base Wide', header=1)
        def _n(s): return _ud.normalize('NFD', str(s).lower()).encode('ascii','ignore').decode()
        col_c = next((c for c in df.columns if any(k in _n(c) for k in
                      ['codigo del centro','servicio de tratamiento'])
                      and 'trabajo' not in _n(c)), None)
        if col_c:
            return sorted(df[col_c].dropna().astype(str).str.strip().unique().tolist())
        return []

def run_paquetes_centros(wide_path, keys_sel=None, progress_cb=None):
    if keys_sel is None:
        keys_sel = list(OUTPUTS.keys())

    centros = _detectar_centros(wide_path)
    if not centros:
        raise ValueError('No se detectaron centros en la base Wide IRT.')

    n_centros = len(centros)
    zip_buf = BytesIO()

    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for i, centro in enumerate(centros):
            slug = _slug(centro)
            carpeta = f'{slug}/'
            if progress_cb: progress_cb(i, n_centros, centro)

            for key in keys_sel:
                out_fname, _ = OUTPUTS[key]
                base_name = out_fname.rsplit('.', 1)[0]
                ext       = out_fname.rsplit('.', 1)[1]
                try:
                    buf, _, _ = run_script(key, wide_path, filtro_centro=centro)
                    zf.writestr(f'{carpeta}{base_name}_{slug}.{ext}', buf.getvalue())
                except Exception as e:
                    zf.writestr(f'{carpeta}ERROR_{key}_{slug}.txt',
                                f'Error generando {out_fname}: {e}')

    if progress_cb: progress_cb(n_centros, n_centros, 'listo')
    zip_buf.seek(0)
    return zip_buf
