"""
pipeline/runner.py — Runner para la app QALAT IRT
Usa exec() en el mismo proceso (igual que el TOP) para mayor compatibilidad.
"""
import sys, os, re, tempfile, shutil, types, traceback as _tb, builtins
from io import BytesIO
from pathlib import Path

PIPELINE_DIR = Path(__file__).parent

OUTPUTS = {
    'caract_excel': ('IRT_Caracterizacion_Ingreso.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
    'seg_excel':    ('IRT_Seguimiento.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
    'word_caract':  ('IRT_Informe_Caracterizacion.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
    'word_seg':     ('IRT_Informe_Seguimiento.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
    'pptx_caract':  ('IRT_Presentacion_Caracterizacion.pptx','application/vnd.openxmlformats-officedocument.presentationml.presentation'),
    'pptx_seg':     ('IRT_Presentacion_Seguimiento.pptx','application/vnd.openxmlformats-officedocument.presentationml.presentation'),
}

SCRIPT_FILES = {
    'caract_excel': 'caract_excel.py',
    'seg_excel':    'seg_excel.py',
    'word_caract':  'word_caract.py',
    'word_seg':     'word_seg.py',
    'pptx_caract':  'pptx_caract.py',
    'pptx_seg':     'pptx_seg.py',
}


def _exec_script(script_key, wide_path, out_path, filtro_centro=None):
    """
    Carga el script IRT, inyecta rutas en el código fuente y ejecuta
    con exec() en el mismo proceso. Maneja dos estilos de declaración:
      - Excel: INPUT_FILE = auto_archivo_wide() / OUTPUT_FILE = '/home/claude/...'
      - Word:  INPUT_FILE = None  # runner inyecta la ruta real
    """
    src = open(str(PIPELINE_DIR / SCRIPT_FILES[script_key]), encoding='utf-8').read()

    wide_esc = wide_path.replace('\\', '/')
    out_esc  = out_path.replace('\\', '/')

    # ── 1. INPUT_FILE — estilo Excel (auto_archivo_wide) ─────────────────────
    src = re.sub(
        r'INPUT_FILE\s*=\s*auto_archivo_wide\(\)',
        f'INPUT_FILE = r"""{wide_esc}"""',
        src
    )
    # ── 2. INPUT_FILE — estilo Word (None + comentario runner) ───────────────
    src = re.sub(
        r'INPUT_FILE\s*=\s*None\s*#.*runner.*',
        f'INPUT_FILE = r"""{wide_esc}"""',
        src
    )

    # ── 3. OUTPUT_FILE — estilo Excel (path literal /home/claude/...) ────────
    src = re.sub(
        r"OUTPUT_FILE\s*=\s*'/home/claude/[^']*\.(?:xlsx|docx|pptx)'",
        f'OUTPUT_FILE = r"""{out_esc}"""',
        src
    )
    src = re.sub(
        r'OUTPUT_FILE\s*=\s*"/home/claude/[^"]*.(?:xlsx|docx|pptx)"',
        f'OUTPUT_FILE = r"""{out_esc}"""',
        src
    )
    # ── 4. OUTPUT_FILE — estilo Word (None + comentario runner) ──────────────
    src = re.sub(
        r'OUTPUT_FILE\s*=\s*None\s*#.*runner.*',
        f'OUTPUT_FILE = r"""{out_esc}"""',
        src
    )
    # ── 5. OUTPUT_FILE — overrides con f-string dentro de if FILTRO_CENTRO ───
    #     Los neutralizamos para que no sobreescriban la ruta correcta
    src = re.sub(
        r"OUTPUT_FILE\s*=\s*f'/home/claude/[^']*'",
        f'OUTPUT_FILE = r"""{out_esc}"""',
        src
    )
    src = re.sub(
        r'OUTPUT_FILE\s*=\s*f"/home/claude/[^"]*"',
        f'OUTPUT_FILE = r"""{out_esc}"""',
        src
    )

    # ── 6. FILTRO_CENTRO ──────────────────────────────────────────────────────
    if filtro_centro:
        centro_esc = str(filtro_centro).replace('"', '\\"')
        # Estilo Word: FILTRO_CENTRO = None   # runner inyecta
        src = re.sub(
            r'FILTRO_CENTRO\s*=\s*None\s*#.*runner.*',
            f'FILTRO_CENTRO = "{centro_esc}"',
            src
        )
        # Estilo Excel: FILTRO_CENTRO = None (línea sola)
        src = re.sub(
            r'^FILTRO_CENTRO\s*=\s*None\s*$',
            f'FILTRO_CENTRO = "{centro_esc}"',
            src, flags=re.MULTILINE
        )
    # Si filtro_centro es None, no tocamos nada (queda None)

    # ── 7. Redirigir auto_archivo_wide() para que no busque en disco ──────────
    src = re.sub(
        r'def auto_archivo_wide\(\):.*?(?=\ndef |\nprint\(|\n[A-Z_]|\Z)',
        'def auto_archivo_wide():\n    return INPUT_FILE\n',
        src, flags=re.DOTALL
    )

    # ── 8. Parchear if __name__ == '__main__': → ejecutar siempre ─────────────
    src = re.sub(
        r"if\s+__name__\s*==\s*['\"]__main__['\"]\s*:",
        'if True:  # runner: ejecutar siempre',
        src
    )

    # ── 9. Parchear accesos directos a 'Tiene_IRT3' que crashean si no existe ─
    # caract_excel: línea de print con .sum()
    src = src.replace(
        '(df["Tiene_IRT3"]=="Sí").sum()',
        '(df["Tiene_IRT3"].eq("Sí").sum() if "Tiene_IRT3" in df.columns else 0)'
    )
    # seg_excel: mask3 = df['Tiene_IRT3'] == 'Sí'
    src = re.sub(
        r"^(\s*)mask3\s*=\s*df\['Tiene_IRT3'\]\s*==\s*'Sí'\s*$",
        r"\1mask3 = df['Tiene_IRT3'].eq('Sí') if 'Tiene_IRT3' in df.columns else __import__('pandas').Series([False]*len(df), index=df.index)",
        src, flags=re.MULTILINE
    )

    # ── 10. Redirigir rutas /home/claude/ temporales → /tmp/ (para pptx) ──────
    #     Los scripts pptx escriben .json y .js en /home/claude/ que no existe
    #     en Streamlit Cloud. Se redirigen a /tmp/ que sí existe.
    src = re.sub(
        r"'/home/claude/(_irt[^']*)'",
        r"'/tmp/\1'",
        src
    )
    src = re.sub(
        r'"/home/claude/(_irt[^"]*)"',
        r'"/tmp/\1"',
        src
    )

    # ── 8. Ejecutar con exec() ────────────────────────────────────────────────
    mod = types.ModuleType(f'_qmod_irt_{script_key}')
    mod.__file__ = str(PIPELINE_DIR / SCRIPT_FILES[script_key])
    mod.__dict__['__builtins__'] = builtins

    try:
        exec(compile(src, f'<{script_key}>', 'exec'), mod.__dict__)
    except SystemExit as e:
        # SystemExit(0) = fin normal (ej: sin datos IRT2)
        if e.code and e.code != 0:
            raise RuntimeError(f'El script terminó con código de error: {e.code}')
    except Exception:
        raise RuntimeError(
            f'Error ejecutando {script_key}:\n{_tb.format_exc()}'
        )


def run_script(script_key, wide_path, filtro_centro=None):
    out_filename, mimetype = OUTPUTS[script_key]
    if filtro_centro:
        base, ext = out_filename.rsplit('.', 1)
        out_filename = f'{base}_{filtro_centro}.{ext}'

    suffix = '.' + out_filename.rsplit('.', 1)[1]
    fd, out_path = tempfile.mkstemp(suffix=suffix, prefix='qalat_irt_out_')
    os.close(fd)

    try:
        _exec_script(script_key, wide_path, out_path, filtro_centro)

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
