#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   SCRIPT_IRT_Universal_PPTX_Seguimiento.py  —  v1.1                       ║
║   Genera presentación PowerPoint de seguimiento IRT1 → IRT2               ║
║   8 slides · Compatible con cualquier país IRT                             ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  CÓMO USAR:                                                                 ║
║  1. Sube este script + la base Wide IRT                                    ║
║  2. Escribe: "Ejecuta el PPTX Seguimiento IRT"                            ║
║                                                                             ║
║  SLIDES:                                                                    ║
║    1. Portada                                                               ║
║    2. Antecedentes generales (sexo + instrumentos)                        ║
║    3. Sustancia principal IRT1 vs IRT2 (tortas)                           ║
║    4. Días de consumo + Cambio en consumo (barras apiladas)               ║
║    5. Urgencias + Accidentes (2 gráficos lado a lado)                     ║
║    6. Salud psicológica y física + Satisfacción de vida                   ║
║    7. Transgresión total + por tipo                                        ║
║    8. Relaciones interpersonales                                           ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""
import glob, os, unicodedata

def _norm(s):
    return unicodedata.normalize('NFD', str(s).lower()).encode('ascii','ignore').decode()

# ── Detección de país ─────────────────────────────────────────────────────────
_PAISES = {
    'republica_dominicana':'República Dominicana', 'repdomini':'República Dominicana',
    'dominicana':'República Dominicana', 'honduras':'Honduras',
    'panama':'Panamá', 'panam':'Panamá', 'el_salvador':'El Salvador',
    'salvador':'El Salvador', 'mexico':'México', 'mexic':'México',
    'ecuador':'Ecuador', 'peru':'Perú', 'argentina':'Argentina',
    'colombia':'Colombia', 'chile':'Chile', 'bolivia':'Bolivia',
    'paraguay':'Paraguay', 'uruguay':'Uruguay', 'venezuela':'Venezuela',
    'guatemala':'Guatemala', 'costa_rica':'Costa Rica',
    'costarica':'Costa Rica', 'nicaragua':'Nicaragua',
}
def _extraer_pais(filename):
    fn = _norm(str(filename).replace('.','_'))
    for key, nombre in _PAISES.items():
        if key in fn: return nombre
    return None

def _detectar_pais(wide_file):
    import pandas as _pd
    try:
        rs = _pd.read_excel(wide_file, sheet_name='Resumen', header=None)
        for _, row in rs.iterrows():
            for v in row.tolist():
                p = _extraer_pais(str(v))
                if p: return p
    except: pass
    return _extraer_pais(os.path.basename(wide_file))

def auto_archivo_wide():
    candidatos = (
        glob.glob('/mnt/user-data/uploads/*IRT*Wide*.xlsx') +
        glob.glob('/mnt/user-data/uploads/*Wide*IRT*.xlsx') +
        glob.glob('/mnt/user-data/uploads/IRT_Base*.xlsx') +
        glob.glob('/mnt/user-data/outputs/IRT_Base_Wide*.xlsx') +
        glob.glob('/home/claude/IRT_Base_Wide.xlsx'))
    if not candidatos:
        raise FileNotFoundError('⚠  No se encontró la base Wide IRT.')
    uploads = [f for f in candidatos if 'uploads' in f]
    elegido = uploads[0] if uploads else max(candidatos, key=os.path.getsize)
    print(f'  → Base Wide: {os.path.basename(elegido)}')
    return elegido

# ══════════════════════════════════════════════════════════════════════════════
print('=' * 60)
print('  SCRIPT_IRT_Universal_PPTX_Seguimiento  v1.1')
print('=' * 60)

INPUT_FILE  = auto_archivo_wide()
OUTPUT_FILE = '/home/claude/IRT_Presentacion_Seguimiento.pptx'

# ── FILTRO OPCIONAL POR CENTRO ────────────────────────────────────────────────
# Dejar en None para procesar TODOS los centros.
# Ejemplos:
#   FILTRO_CENTRO = None         ← todos los centros
#   FILTRO_CENTRO = "HCHN01"     ← solo ese centro
FILTRO_CENTRO = None
# ─────────────────────────────────────────────────────────────────────────────

import pandas as pd, numpy as np, json, subprocess, sys, warnings
warnings.filterwarnings('ignore')

# ── Carga ─────────────────────────────────────────────────────────────────────
df = pd.read_excel(INPUT_FILE, sheet_name='Base Wide', header=1)
df.columns = [str(c) for c in df.columns]
cols = df.columns.tolist()

_col_centro = next((c for c in cols if any(x in _norm(c) for x in
                    ['codigo del centro', 'servicio de tratamiento',
                     'centro/ servicio', 'codigo centro'])), None)

if FILTRO_CENTRO and _col_centro:
    n_antes = len(df)
    df = df[df[_col_centro].astype(str).str.strip() == FILTRO_CENTRO].copy()
    df = df.reset_index(drop=True)
    print(f'\n⚑  Filtro: "{FILTRO_CENTRO}"  ({n_antes} → {len(df)} pacientes)')
    OUTPUT_FILE = f'/home/claude/IRT_Presentacion_Seguimiento_{FILTRO_CENTRO}.pptx'

N_total = len(df)
mask2   = df['Tiene_IRT2'] == 'Sí'
mask3   = df['Tiene_IRT3'] == 'Sí'
N_irt2  = int(mask2.sum())
N_irt3  = int(mask3.sum())
TIENE_IRT3 = N_irt3 > 0

print(f'\n→ {N_total} pacientes | IRT2: {N_irt2} | IRT3: {N_irt3}')
if N_irt2 == 0:
    raise SystemExit('⚠  Sin IRT2. No se puede generar la presentación.')

# ── País y servicio ────────────────────────────────────────────────────────────
PAIS = _detectar_pais(INPUT_FILE)
if FILTRO_CENTRO:
    NOMBRE_SERVICIO = f'{PAIS}  —  Centro {FILTRO_CENTRO}' if PAIS else f'Centro {FILTRO_CENTRO}'
else:
    NOMBRE_SERVICIO = PAIS if PAIS else 'Servicio de Tratamiento'

# ── Período ───────────────────────────────────────────────────────────────────
MESES = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',
         7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}
fecha_col = next((c for c in cols if 'fecha de administracion' in _norm(c)), None)
PERIODO = 'Período no determinado'
if fecha_col:
    fch = pd.to_datetime(df[fecha_col], errors='coerce').dropna()
    anio = pd.Timestamp.now().year
    fch  = fch[(fch.dt.year >= anio-10) & (fch.dt.year <= anio+1)]
    if len(fch):
        f0,f1 = fch.min(), fch.max()
        PERIODO = (f'{MESES[f0.month]} {f0.year}'
                   if f0.year==f1.year and f0.month==f1.month
                   else f'{MESES[f0.month]}–{MESES[f1.month]} {f0.year}'
                   if f0.year==f1.year
                   else f'{MESES[f0.month]} {f0.year} – {MESES[f1.month]} {f1.year}')

print(f'  Servicio: {NOMBRE_SERVICIO} | Período: {PERIODO}')

# ══════════════════════════════════════════════════════════════════════════════
# DETECCIÓN DE COLUMNAS
# ══════════════════════════════════════════════════════════════════════════════
def col_sfx(kws, sfx):
    for c in cols:
        if not c.endswith(sfx): continue
        if all(_norm(k) in _norm(c) for k in kws): return c
    return None

COL_FN   = next((c for c in cols if 'fecha de nacimiento' in _norm(c)), None)
COL_SEXO = next((c for c in cols if _norm(c) in ['sexo','género','genero']), None)

SUST_NOMBRES = {
    'Alcohol':['alcohol'], 'Marihuana':['marihuana','cannabis'],
    'Heroína':['heroina'], 'Cocaína':['cocain'],
    'Fentanilo':['fentanil'], 'Inhalables':['inhalab'],
    'Metanfetamina':['metanfet','cristal'], 'Crack':['crack'],
    'Pasta Base':['pasta base','pasta'], 'Sedantes':['sedant','benzod'],
    'Otra sustancia':['otra sust'],
}
SUST_TOTAL = {}
for sust, kws in SUST_NOMBRES.items():
    entry = {}
    for sfx in ['_IRT1','_IRT2','_IRT3']:
        for c in cols:
            if not c.endswith(sfx): continue
            nc = _norm(c)
            if any(_norm(k) in nc for k in kws) and ('total' in nc or '(0-28)' in nc):
                entry[sfx] = c; break
    if entry: SUST_TOTAL[sust] = entry
SUST_ACTIVAS = list(SUST_TOTAL.keys())

COL_SP   = {sfx: col_sfx(['sustancia','principal'], sfx) for sfx in ['_IRT1','_IRT2','_IRT3']}
COL_SPSI = {sfx: col_sfx(['salud','psicol'], sfx)         for sfx in ['_IRT1','_IRT2','_IRT3']}
COL_SFIS = {sfx: col_sfx(['salud','fis'],    sfx)         for sfx in ['_IRT1','_IRT2','_IRT3']}
COL_URG  = {sfx: next((c for c in cols if c.endswith(sfx) and '5)' in c and
                        any(k in c.lower() for k in ['urgencia','hospitali','emergencia'])), None)
            for sfx in ['_IRT1','_IRT2','_IRT3']}
COL_ACC  = {sfx: next((c for c in cols if c.endswith(sfx) and '6)' in c and
                        'accidente' in c.lower()), None)
            for sfx in ['_IRT1','_IRT2','_IRT3']}
TRANS_DEF = {'Robo / Hurto':'robo','Venta de sustancias':'venta',
             'Violencia':'violencia','VIF':'intraf','Detenido':'detenido'}
TRANS_COLS = {}
for nombre, kw in TRANS_DEF.items():
    entry = {sfx: next((c for c in cols if c.endswith(sfx) and kw in c.lower()), None)
             for sfx in ['_IRT1','_IRT2']}
    if any(entry.values()): TRANS_COLS[nombre] = entry

REL_MAP = {'Padre':['padre'],'Madre':['madre'],'Hijos':['hijos','hijo'],
           'Hermanos':['hermanos'],'Pareja':['pareja'],'Amigos':['amigos'],'Otros':['otros']}
REL_COLS = {}
for vin, kws in REL_MAP.items():
    for sfx in ['_IRT1','_IRT2']:
        for c in cols:
            if not c.endswith(sfx): continue
            nc = _norm(c)
            if '14)' not in c and 'relaci' not in nc: continue
            if any(_norm(k) in nc for k in kws):
                REL_COLS.setdefault(vin,{})[sfx]=c; break

SAT_MAP = {
    'Vida general':[['16)'],['satisfac','vida']],
    'Lugar donde vive':[['17)'],['satisfac','lugar']],
    'Situación laboral':[['18)'],['satisfac','labor','educac']],
    'Tiempo libre':[['19)'],['satisfac','tiempo']],
    'Cap. económica':[['20)'],['satisfac','econom']],
}
SAT_COLS = {}
for dim,(nums,kws) in SAT_MAP.items():
    for sfx in ['_IRT1','_IRT2']:
        for c in cols:
            if not c.endswith(sfx): continue
            nc = _norm(c)
            if any(n in c for n in nums) and any(_norm(k) in nc for k in kws):
                SAT_COLS.setdefault(dim,{})[sfx]=c; break

# ══════════════════════════════════════════════════════════════════════════════
# CALCULAR DATOS
# ══════════════════════════════════════════════════════════════════════════════
print('\n→ Calculando datos...')

def norm_sust(s):
    if pd.isna(s) or str(s).strip() in ['0','']: return None
    s = _norm(str(s))
    if any(x in s for x in ['alcohol','cerveza','licor']): return 'Alcohol'
    if any(x in s for x in ['marihu','cannabis','marij']): return 'Marihuana'
    if any(x in s for x in ['crack','piedra','paco']):     return 'Crack'
    if any(x in s for x in ['pasta base','pasta']):        return 'Pasta Base'
    if any(x in s for x in ['cocain','perico']):           return 'Cocaína'
    if any(x in s for x in ['fentanil']):                  return 'Fentanilo'
    if any(x in s for x in ['inhalab']):                   return 'Inhalables'
    if any(x in s for x in ['metanfet','cristal']):        return 'Metanfetamina'
    if any(x in s for x in ['sedant','benzod']):           return 'Sedantes'
    if any(x in s for x in ['heroina','opiod','morfin']):  return 'Heroína'
    if any(x in s for x in ['tabaco','nicot']):            return 'Tabaco'
    return 'Otra sustancia'

def prom(col, msk):
    if col is None: return None
    v = pd.to_numeric(df.loc[msk,col], errors='coerce').dropna()
    return round(float(v.mean()),1) if len(v) else None

def sino_pct(col, msk):
    if col is None: return None, None
    v = df.loc[msk,col].dropna().astype(str).str.strip().str.lower()
    nv = len(v)
    nsi = int(v.isin(['sí','si']).sum())
    return nsi, round(nsi/nv*100,1) if nv else 0.0

hoy = pd.Timestamp.now()
pct_seg = round(N_irt2/N_total*100,1)

# Sexo
sexo_data = {}
if COL_SEXO:
    sc = df.loc[mask2, COL_SEXO].value_counts(dropna=True)
    sexo_data = {k:int(v) for k,v in sc.items()}

# Edad
edad_data = {}
if COL_FN:
    edades = ((hoy - pd.to_datetime(df.loc[mask2,COL_FN], errors='coerce')).dt.days/365.25).dropna()
    edad_data = {'mean': round(edades.mean(),1), 'min': int(edades.min()),
                 'max': int(edades.max())} if len(edades) else {}

# Sustancia principal IRT1 vs IRT2
def dist_sp(sfx, msk):
    c = COL_SP.get(sfx)
    if not c: return {}
    return df.loc[msk,c].apply(norm_sust).dropna().value_counts().to_dict()

sp1 = dist_sp('_IRT1', mask2)
sp2 = dist_sp('_IRT2', mask2)

# Días consumo por sustancia IRT1 vs IRT2
dias_data = []
for sust in SUST_ACTIVAS:
    entry = SUST_TOTAL[sust]
    c1 = entry.get('_IRT1'); c2 = entry.get('_IRT2')
    if not c1: continue
    v1c = pd.to_numeric(df.loc[mask2,c1],errors='coerce'); v1c = v1c[v1c>0].dropna()
    v2c = pd.to_numeric(df.loc[mask2,c2],errors='coerce') if c2 else pd.Series(dtype=float)
    v2c = v2c[v2c>0].dropna()
    m1 = round(float(v1c.mean()),1) if len(v1c) else 0.0
    m2 = round(float(v2c.mean()),1) if len(v2c) else 0.0
    if m1 > 0 or m2 > 0:
        dias_data.append({'label':sust,'irt1':m1,'irt2':m2})
dias_data.sort(key=lambda x:-x['irt1'])

# Cambio en consumo (barras apiladas) — solo consumidores IRT1
cambio_data = []
for sust in SUST_ACTIVAS:
    entry = SUST_TOTAL[sust]
    c1 = entry.get('_IRT1'); c2 = entry.get('_IRT2')
    if not c1 or not c2: continue
    sp1col = COL_SP.get('_IRT1')
    if not sp1col: continue
    # Solo personas cuya sustancia principal es esta
    msk_pp = mask2 & (df[sp1col].apply(norm_sust) == sust)
    if not msk_pp.any(): continue
    v1 = pd.to_numeric(df.loc[msk_pp,c1],errors='coerce')
    v2 = pd.to_numeric(df.loc[msk_pp,c2],errors='coerce')
    ok = v1.notna() & v2.notna()
    # También incluir cualquier consumidor en IRT1
    msk_cons = mask2 & (pd.to_numeric(df[c1],errors='coerce').fillna(0) > 0) & pd.to_numeric(df[c2],errors='coerce').notna()
    v1b = pd.to_numeric(df.loc[msk_cons,c1],errors='coerce')
    v2b = pd.to_numeric(df.loc[msk_cons,c2],errors='coerce')
    ok2 = v1b.notna() & v2b.notna()
    n_ok = int(ok2.sum())
    if n_ok == 0: continue
    n_abs = int((v2b[ok2]==0).sum())
    n_dis = int(((v2b[ok2]>0)&(v2b[ok2]<v1b[ok2])).sum())
    n_sc  = int((v2b[ok2]==v1b[ok2]).sum())
    n_emp = int((v2b[ok2]>v1b[ok2]).sum())
    pct = lambda n: round(n/n_ok*100,1)
    cambio_data.append({'label':sust,'n':n_ok,
        'abs':pct(n_abs),'dis':pct(n_dis),'sin':pct(n_sc),'emp':pct(n_emp),
        'combo':round((n_abs+n_dis)/n_ok*100,1)})

# Urgencias y accidentes
n_urg1, pct_urg1 = sino_pct(COL_URG.get('_IRT1'), mask2)
n_urg2, pct_urg2 = sino_pct(COL_URG.get('_IRT2'), mask2)
n_acc1, pct_acc1 = sino_pct(COL_ACC.get('_IRT1'), mask2)
n_acc2, pct_acc2 = sino_pct(COL_ACC.get('_IRT2'), mask2)

# Salud
salud_data = []
for nombre, cd in [('Salud Psicológica (0–10)', COL_SPSI),
                   ('Salud Física (0–10)',       COL_SFIS)]:
    m1 = prom(cd.get('_IRT1'), mask2)
    m2 = prom(cd.get('_IRT2'), mask2)
    if m1 is not None or m2 is not None:
        salud_data.append({'label':nombre,'irt1':m1 or 0,'irt2':m2 or 0})

# Satisfacción
sat_data = []
for dim, cd in SAT_COLS.items():
    m1 = prom(cd.get('_IRT1'), mask2)
    m2 = prom(cd.get('_IRT2'), mask2)
    if m1 is not None or m2 is not None:
        sat_data.append({'label':dim,'irt1':m1 or 0,'irt2':m2 or 0})

# Transgresión
def mask_trans(sfx, msk):
    cols_tr = [d.get(sfx) for d in TRANS_COLS.values() if d.get(sfx)]
    if not cols_tr: return pd.Series(False, index=df.index[msk])
    return pd.concat([pd.to_numeric(df.loc[msk,c],errors='coerce').fillna(0)>0
                      for c in cols_tr], axis=1).any(axis=1)

n_tr1 = int(mask_trans('_IRT1',mask2).sum())
n_tr2 = int(mask_trans('_IRT2',mask2).sum())
pct_tr1 = round(n_tr1/N_irt2*100,1); pct_tr2 = round(n_tr2/N_irt2*100,1)

tipos_tr = []
for tipo, cd in TRANS_COLS.items():
    c1=cd.get('_IRT1'); c2=cd.get('_IRT2')
    v1 = pd.to_numeric(df.loc[mask2,c1],errors='coerce').fillna(0) if c1 else pd.Series([0]*N_irt2)
    v2 = pd.to_numeric(df.loc[mask2,c2],errors='coerce').fillna(0) if c2 else pd.Series([0]*N_irt2)
    p1 = round((v1>0).sum()/N_irt2*100,1); p2 = round((v2>0).sum()/N_irt2*100,1)
    if p1>0 or p2>0:
        tipos_tr.append({'label':tipo,'irt1':p1,'irt2':p2})

# Relaciones
CATS_REL = ['Excelente','Buena','Ni buena ni mala','Mala','Muy mala']
rel_data = []
for vin, cd in REL_COLS.items():
    row_rel = {'label': vin}
    for sfx, lbl in [('_IRT1','irt1'),('_IRT2','irt2')]:
        c = cd.get(sfx)
        if not c: row_rel[f'pos_{lbl}']=0; row_rel[f'neg_{lbl}']=0; continue
        vals = df.loc[mask2,c].dropna()
        vals = vals[vals.astype(str).str.lower()!='no aplica']
        nv = len(vals)
        if nv == 0: row_rel[f'pos_{lbl}']=0; row_rel[f'neg_{lbl}']=0; continue
        pos = int((vals.astype(str).isin(['Excelente','Buena'])).sum())
        neg = int((vals.astype(str).isin(['Mala','Muy mala'])).sum())
        row_rel[f'pos_{lbl}'] = round(pos/nv*100,1)
        row_rel[f'neg_{lbl}'] = round(neg/nv*100,1)
    rel_data.append(row_rel)

print('  ✓ Datos calculados')
print(f'  Sustancias: {SUST_ACTIVAS}')
print(f'  Cambio: {len(cambio_data)} sustancias | Trans. IRT1: {pct_tr1}% → IRT2: {pct_tr2}%')

# ══════════════════════════════════════════════════════════════════════════════
# JSON INTERMEDIO
# ══════════════════════════════════════════════════════════════════════════════
data = {
    'meta': {
        'servicio': NOMBRE_SERVICIO,
        'periodo':  PERIODO,
        'N_total':  N_total,
        'N_irt2':   N_irt2,
        'N_irt3':   N_irt3,
        'pct_seg':  pct_seg,
    },
    'sexo':    [{'label':k,'n':v,'pct':round(v/N_irt2*100,1)} for k,v in sexo_data.items()],
    'edad':    edad_data,
    'sp1':     [{'label':k,'n':v,'pct':round(v/N_irt2*100,1)} for k,v in sp1.items()],
    'sp2':     [{'label':k,'n':v,'pct':round(v/N_irt2*100,1)} for k,v in sp2.items()],
    'dias':    dias_data,
    'cambio':  cambio_data,
    'urgencias': {
        'n1':   n_urg1, 'pct1': pct_urg1,
        'n2':   n_urg2, 'pct2': pct_urg2,
    },
    'accidentes': {
        'n1':   n_acc1, 'pct1': pct_acc1,
        'n2':   n_acc2, 'pct2': pct_acc2,
    },
    'salud':   salud_data,
    'sat':     sat_data,
    'transTotal': {'irt1': pct_tr1, 'irt2': pct_tr2, 'n1': n_tr1, 'n2': n_tr2},
    'transTipos': tipos_tr,
    'rel':     rel_data,
}

json_path = '/home/claude/_irt_data.json'
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

# ══════════════════════════════════════════════════════════════════════════════
# NODE.JS — construye el PowerPoint con pptxgenjs
# ══════════════════════════════════════════════════════════════════════════════
JS_CODE = r"""
const fs      = require('fs');
const pptxgen = require('pptxgenjs');

const data   = JSON.parse(fs.readFileSync('/home/claude/_irt_data.json', 'utf8'));
const OUTPUT = '""" + OUTPUT_FILE + r"""';

// ── Paleta ──────────────────────────────────────────────────────────────────
const C_DARK   = '1F3864', C_MID  = '2E75B6', C_LIGHT = 'BDD7EE';
const C_IRT1   = '1F3864', C_IRT2 = '00B0F0', C_IRT3  = '70AD47';
const C_GRAY   = '595959', C_WHITE = 'FFFFFF', C_TITLE = '0070C0';
const C_ABS    = '1F3864', C_DIS  = '375623', C_SC    = 'BFBFBF', C_EMP = 'C00000';
const PIE_COLS = ['2E75B6','1F3864','00B0F0','9DC3E6','70AD47','4472C4',
                  'BFBFBF','C00000','ED7D31','FFC000','7030A0','538135'];

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

// Header estándar para slides de contenido
function hdr(sl, txt) {
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:10,h:0.70,
    fill:{color:C_DARK},line:{color:C_DARK}});
  sl.addShape(pres.shapes.RECTANGLE, {x:5.2,y:0,w:4.8,h:0.70,
    fill:{color:C_MID,transparency:40},line:{color:C_MID,transparency:40}});
  sl.addText(txt, {x:0.25,y:0,w:9.5,h:0.70,
    fontSize:20,bold:true,color:C_WHITE,fontFace:'Calibri',valign:'middle'});
}

const TITULO = `Resultados: Ingreso y 3 meses${data.meta.N_irt3 > 0 ? ' y 6 meses' : ''}`;

// ── SLIDE 1: PORTADA ───────────────────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:4.0,h:5.625,
    fill:{color:C_DARK},line:{color:C_DARK}});
  sl.addShape(pres.shapes.RECTANGLE, {x:3.1,y:0,w:1.5,h:5.625,
    fill:{color:C_MID,transparency:60},line:{color:C_MID,transparency:60}});
  sl.addText('Resultados', {x:0.25,y:1.6,w:3.2,h:0.7,
    fontSize:24,bold:true,color:C_WHITE,fontFace:'Calibri'});
  sl.addText('Monitoreo IRT', {x:0.25,y:2.35,w:3.2,h:0.55,
    fontSize:14,color:C_LIGHT,fontFace:'Calibri'});
  sl.addText([
    {text:'IRT 1 - IRT 2', options:{breakLine:true}},
    {text:'Ingreso y Seguimiento'}
  ], {x:4.6,y:1.55,w:5.1,h:1.4,
    fontSize:30,bold:true,color:C_GRAY,fontFace:'Calibri',align:'center',valign:'middle'});
  sl.addText(data.meta.servicio.toUpperCase(), {x:4.6,y:3.1,w:5.1,h:0.45,
    fontSize:17,bold:true,color:C_MID,fontFace:'Calibri',align:'center'});
  sl.addText(data.meta.periodo, {x:4.6,y:3.6,w:5.1,h:0.35,
    fontSize:13,bold:true,color:C_MID,fontFace:'Calibri',align:'center'});
  sl.addText(
    `N = ${data.meta.N_irt2} personas con seguimiento (${data.meta.pct_seg}% del total)`,
    {x:4.6,y:4.1,w:5.1,h:0.35,fontSize:11,color:'888888',fontFace:'Calibri',align:'center'});
}

// ── SLIDE 2: ANTECEDENTES GENERALES ───────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addShape(pres.shapes.LINE, {x:4.95,y:0.76,w:0,h:4.85,
    line:{color:'D9D9D9',width:1}});

  // Izquierda: tabla instrumentos + sexo
  sl.addText('Instrumentos contestados',
    {x:0.25,y:0.82,w:4.5,h:0.38,
     fontSize:13,bold:true,color:C_GRAY,fontFace:'Calibri'});
  const tInst = [
    [{text:'Instrumento',options:{bold:true,fontSize:10,color:'FFFFFF',align:'center',fill:{color:C_DARK}}},
     {text:'n',          options:{bold:true,fontSize:10,color:'FFFFFF',align:'center',fill:{color:C_DARK}}}],
    [{text:'IRT 1 — Ingreso',        options:{fontSize:10,color:'363636'}},
     {text:`${data.meta.N_total}`,   options:{fontSize:11,bold:true,color:C_DARK,align:'center'}}],
    [{text:'IRT 2 — Seguimiento 3m', options:{fontSize:10,color:'363636'}},
     {text:`${data.meta.N_irt2}`,    options:{fontSize:11,bold:true,color:C_IRT2,align:'center'}}],
  ];
  if (data.meta.N_irt3 > 0) {
    tInst.push([
      {text:'IRT 3 — Seguimiento 6m', options:{fontSize:10,color:'363636'}},
      {text:`${data.meta.N_irt3}`,    options:{fontSize:11,bold:true,color:C_IRT3,align:'center'}}
    ]);
  }
  sl.addTable(tInst, {x:0.25,y:1.28,w:4.5,h:tInst.length*0.45,
    border:{pt:0.5,color:C_LIGHT},rowH:0.42,
    colW:[3.2,1.1]});

  // Distribución por sexo
  if (data.sexo.length > 0) {
    sl.addText('Distribución por sexo',
      {x:0.25,y:2.65,w:4.5,h:0.35,
       fontSize:12,bold:true,color:C_GRAY,fontFace:'Calibri'});
    sl.addChart(pres.charts.PIE, [{
      name:'Sexo',
      labels: data.sexo.map(s => s.label),
      values: data.sexo.map(s => s.n),
    }], {
      x:0.5,y:3.05,w:4.0,h:2.4,
      showPercent:true,showLabel:false,showLegend:true,legendPos:'b',legendFontSize:10,
      dataLabelFontSize:12,chartColors:PIE_COLS.slice(0,data.sexo.length),
      chartArea:{fill:{color:'FFFFFF'}},
      dataLabelColor:C_WHITE,
    });
  }

  // Derecha: edad
  sl.addText('Distribución por rango de edad',
    {x:5.15,y:0.82,w:4.6,h:0.38,
     fontSize:13,bold:true,color:C_GRAY,fontFace:'Calibri'});
  if (data.edad && data.edad.mean) {
    sl.addText(`Edad promedio: ${data.edad.mean} años  (mín: ${data.edad.min} – máx: ${data.edad.max})`,
      {x:5.15,y:1.22,w:4.6,h:0.30,
       fontSize:9.5,color:C_GRAY,fontFace:'Calibri',italic:true});
  }
  // Tabla de edades (si no hay barras, al menos mostrar tabla)
  const rangesData = [
    {label:'< 18',  cat:'<18'},  {label:'18–25',cat:'18–25'},
    {label:'26–35', cat:'26–35'},{label:'36–45',cat:'36–45'},
    {label:'46–55', cat:'46–55'},{label:'56+',  cat:'56+'},
  ];
  sl.addText(`N seguimiento = ${data.meta.N_irt2}`,
    {x:5.15,y:4.9,w:4.6,h:0.30,
     fontSize:9,color:C_GRAY,fontFace:'Calibri',italic:true,align:'right'});
}

// ── SLIDE 3: SUSTANCIA PRINCIPAL IRT1 vs IRT2 ─────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addText('CONSUMO: SUSTANCIA PRINCIPAL',
    {x:0.25,y:0.78,w:9.5,h:0.38,
     fontSize:14,bold:true,color:C_TITLE,fontFace:'Calibri',align:'center'});
  sl.addShape(pres.shapes.LINE, {x:4.95,y:1.18,w:0,h:4.42,
    line:{color:'D9D9D9',width:1}});

  // IRT1
  sl.addText('IRT 1 — Ingreso', {x:0.25,y:1.22,w:4.5,h:0.35,
    fontSize:12,bold:true,color:C_IRT1,fontFace:'Calibri',align:'center'});
  if (data.sp1.length > 0) {
    sl.addChart(pres.charts.PIE, [{
      name:'Sustancia IRT1',
      labels: data.sp1.map(s=>s.label),
      values: data.sp1.map(s=>s.n),
    }], {
      x:0.3,y:1.55,w:4.5,h:3.8,
      showPercent:true,showLabel:false,showLegend:true,legendPos:'b',legendFontSize:9.5,
      dataLabelFontSize:11,chartColors:PIE_COLS.slice(0,data.sp1.length),
      chartArea:{fill:{color:'FFFFFF'}},dataLabelColor:C_WHITE,
    });
  }

  // IRT2
  sl.addText('IRT 2 — Seguimiento 3 meses', {x:5.2,y:1.22,w:4.5,h:0.35,
    fontSize:12,bold:true,color:C_IRT2,fontFace:'Calibri',align:'center'});
  if (data.sp2.length > 0) {
    sl.addChart(pres.charts.PIE, [{
      name:'Sustancia IRT2',
      labels: data.sp2.map(s=>s.label),
      values: data.sp2.map(s=>s.n),
    }], {
      x:5.2,y:1.55,w:4.5,h:3.8,
      showPercent:true,showLabel:false,showLegend:true,legendPos:'b',legendFontSize:9.5,
      dataLabelFontSize:11,chartColors:PIE_COLS.slice(0,data.sp2.length),
      chartArea:{fill:{color:'FFFFFF'}},dataLabelColor:C_WHITE,
    });
  }
}

// ── SLIDE 4: DÍAS CONSUMO + CAMBIO EN CONSUMO ─────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addShape(pres.shapes.LINE, {x:4.95,y:0.76,w:0,h:4.85,
    line:{color:'D9D9D9',width:1}});

  // Izquierda: días consumo
  sl.addText('PROMEDIO DE DÍAS DE CONSUMO\nIRT 1 (Ingreso) vs IRT 2 (Seguimiento)',
    {x:0.25,y:0.80,w:4.5,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  if (data.dias.length > 0) {
    const labs = data.dias.map(d=>d.label);
    sl.addChart(pres.charts.BAR, [
      {name:'Ingreso (IRT 1)',     labels:labs, values:data.dias.map(d=>d.irt1)},
      {name:'Seguimiento (IRT 2)', labels:labs, values:data.dias.map(d=>d.irt2)},
    ], {
      x:0.2,y:1.5,w:4.5,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_IRT1,C_IRT2],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFontSize:10,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:10,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:28,valAxisMinVal:0,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
      showLegend:true,legendPos:'b',legendFontSize:10,
    });
  }

  // Derecha: cambio en consumo (barras apiladas 100%)
  sl.addText('CAMBIO EN EL CONSUMO\nIngreso → Seguimiento',
    {x:5.15,y:0.80,w:4.6,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  if (data.cambio.length > 0) {
    const labs = data.cambio.map(d=>d.label);
    sl.addChart(pres.charts.BAR, [
      {name:'Abstinencia', labels:labs, values:data.cambio.map(d=>d.abs)},
      {name:'Disminuyó',   labels:labs, values:data.cambio.map(d=>d.dis)},
      {name:'Sin cambio',  labels:labs, values:data.cambio.map(d=>d.sin)},
      {name:'Empeoró',     labels:labs, values:data.cambio.map(d=>d.emp)},
    ], {
      x:5.15,y:1.5,w:4.6,h:3.0,barDir:'col',barGrouping:'percentStacked',
      chartColors:[C_ABS,C_DIS,C_SC,C_EMP],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:9,dataLabelColor:C_WHITE,
      catAxisLabelColor:'363636',catAxisLabelFontSize:9.5,
      valAxisLabelColor:'595959',valAxisLabelFontSize:8,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
      showLegend:true,legendPos:'b',legendFontSize:9,
    });
    // Mini-tabla % Abs + Disminuyó
    sl.addText('Abstinencia o reducción',
      {x:5.15,y:4.58,w:4.6,h:0.28,fontSize:9,color:C_GRAY,fontFace:'Calibri',align:'center',italic:true});
    const cw = 4.4 / data.cambio.length;
    sl.addTable([
      data.cambio.map(d=>({text:d.label,    options:{bold:false,fontSize:8.5, color:'363636',align:'center'}})),
      data.cambio.map(d=>({text:`${d.combo}%`,options:{bold:true, fontSize:13,  color:C_DARK,  align:'center'}})),
    ], {x:5.25,y:4.88,w:4.4,h:0.7,
      border:{pt:0.5,color:C_LIGHT},fill:{color:'EEF4FB'},
      rowH:0.32, colW:data.cambio.map(()=>cw)});
  }
}

// ── SLIDE 5: URGENCIAS + ACCIDENTES ───────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addShape(pres.shapes.LINE, {x:4.95,y:0.76,w:0,h:4.85,
    line:{color:'D9D9D9',width:1}});

  const U = data.urgencias;
  const A = data.accidentes;

  // Izquierda: Urgencias
  sl.addText('Personas que acudieron a Urgencia o estuvieron\nHospitalizadas debido al consumo de Sustancias',
    {x:0.25,y:0.82,w:4.5,h:0.75,
     fontSize:12,bold:true,color:C_GRAY,fontFace:'Calibri',align:'left'});
  if (U.pct1 !== null && U.pct2 !== null) {
    sl.addChart(pres.charts.BAR, [
      {name:'IRT 1', labels:['Urgencias u hospitalización'], values:[U.pct1]},
      {name:'IRT 2', labels:['Urgencias u hospitalización'], values:[U.pct2]},
    ], {
      x:0.3,y:1.62,w:4.4,h:3.68,barDir:'col',barGrouping:'clustered',
      chartColors:[C_IRT1,C_IRT2],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0.0"%"',dataLabelFontSize:13,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:12,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:Math.max(U.pct1, U.pct2, 5)*1.4,
      valAxisNumFmt:'0"%"',
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
      showLegend:true,legendPos:'b',legendFontSize:11,
    });
  }

  // Derecha: Accidentes
  sl.addText('Personas que tuvieron un Accidente relacionado al\nConsumo de Sustancias',
    {x:5.15,y:0.82,w:4.6,h:0.75,
     fontSize:12,bold:true,color:C_GRAY,fontFace:'Calibri',align:'left'});
  if (A.pct1 !== null && A.pct2 !== null) {
    sl.addChart(pres.charts.BAR, [
      {name:'IRT 1', labels:['Accidentes'], values:[A.pct1]},
      {name:'IRT 2', labels:['Accidentes'], values:[A.pct2]},
    ], {
      x:5.15,y:1.62,w:4.5,h:3.68,barDir:'col',barGrouping:'clustered',
      chartColors:[C_IRT1,C_IRT2],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0.0"%"',dataLabelFontSize:13,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:12,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:Math.max(A.pct1, A.pct2, 5)*1.4,
      valAxisNumFmt:'0"%"',
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
      showLegend:true,legendPos:'b',legendFontSize:11,
    });
  }
}

// ── SLIDE 6: SALUD + SATISFACCIÓN ─────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addShape(pres.shapes.LINE, {x:4.95,y:0.76,w:0,h:4.85,
    line:{color:'D9D9D9',width:1}});

  sl.addText('AUTOPERCEPCIÓN DEL ESTADO DE SALUD\n(escala 0–10)',
    {x:0.25,y:0.80,w:4.5,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  if (data.salud.length > 0) {
    sl.addChart(pres.charts.BAR, [
      {name:'Ingreso (IRT 1)',     labels:data.salud.map(d=>d.label), values:data.salud.map(d=>d.irt1)},
      {name:'Seguimiento (IRT 2)', labels:data.salud.map(d=>d.label), values:data.salud.map(d=>d.irt2)},
    ], {
      x:0.2,y:1.5,w:4.5,h:3.85,barDir:'bar',barGrouping:'clustered',
      chartColors:[C_IRT1,C_IRT2],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:10,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,valAxisMaxVal:10,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
      showLegend:true,legendPos:'b',legendFontSize:10,
    });
  }

  sl.addText('SATISFACCIÓN DE VIDA\n(escala 0–10)',
    {x:5.15,y:0.80,w:4.6,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  if (data.sat.length > 0) {
    sl.addChart(pres.charts.BAR, [
      {name:'Ingreso (IRT 1)',     labels:data.sat.map(d=>d.label), values:data.sat.map(d=>d.irt1)},
      {name:'Seguimiento (IRT 2)', labels:data.sat.map(d=>d.label), values:data.sat.map(d=>d.irt2)},
    ], {
      x:5.15,y:1.5,w:4.6,h:3.85,barDir:'bar',barGrouping:'clustered',
      chartColors:[C_IRT1,C_IRT2],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:10,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,valAxisMaxVal:10,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
      showLegend:true,legendPos:'b',legendFontSize:10,
    });
  }
}

// ── SLIDE 7: TRANSGRESIÓN ─────────────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addShape(pres.shapes.LINE, {x:4.95,y:0.76,w:0,h:4.85,
    line:{color:'D9D9D9',width:1}});
  sl.addText('Personas que cometieron alguna\ntransgresión a la norma social',
    {x:0.25,y:0.80,w:4.5,h:0.65,
     fontSize:13,bold:true,color:C_GRAY,fontFace:'Calibri',align:'left'});
  sl.addText('Distribución por tipo de transgresión',
    {x:5.15,y:0.80,w:4.6,h:0.65,
     fontSize:13,bold:true,color:C_GRAY,fontFace:'Calibri',align:'center'});

  const T = data.transTotal;
  sl.addChart(pres.charts.BAR, [
    {name:'IRT 1', labels:['Ingreso\n(IRT 1)','Seguimiento\n(IRT 2)'], values:[T.irt1, null]},
    {name:'IRT 2', labels:['Ingreso\n(IRT 1)','Seguimiento\n(IRT 2)'], values:[null,   T.irt2]},
  ], {
    x:0.2,y:1.5,w:4.5,h:3.85,barDir:'col',barGrouping:'clustered',
    chartColors:[C_IRT1,C_IRT2],chartArea:{fill:{color:'FFFFFF'}},
    showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:14,
    dataLabelColor:C_WHITE,dataLabelPosition:'inEnd',
    catAxisLabelColor:'363636',catAxisLabelFontSize:12,
    valAxisLabelColor:'595959',valAxisLabelFontSize:9,
    valAxisMaxVal:100,valAxisNumFmt:'0"%"',
    valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
  });

  const tiposFilt = data.transTipos.filter(d => d.irt1 > 0 || d.irt2 > 0);
  if (tiposFilt.length > 0) {
    sl.addChart(pres.charts.BAR, [
      {name:'Ingreso (IRT 1)',     labels:tiposFilt.map(d=>d.label), values:tiposFilt.map(d=>d.irt1)},
      {name:'Seguimiento (IRT 2)', labels:tiposFilt.map(d=>d.label), values:tiposFilt.map(d=>d.irt2)},
    ], {
      x:5.15,y:1.5,w:4.6,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_IRT1,C_IRT2],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:9,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:Math.max(...tiposFilt.map(d=>Math.max(d.irt1,d.irt2)))*1.4 || 20,
      valAxisNumFmt:'0"%"',
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
      showLegend:true,legendPos:'b',legendFontSize:10,
    });
  }
}

// ── SLIDE 8: RELACIONES INTERPERSONALES ───────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addText('CALIDAD DE LAS RELACIONES INTERPERSONALES',
    {x:0.25,y:0.78,w:9.5,h:0.38,
     fontSize:14,bold:true,color:C_TITLE,fontFace:'Calibri',align:'center'});
  sl.addText('% de relaciones positivas (Excelente + Buena)  ·  IRT 1 (Ingreso) vs IRT 2 (Seguimiento)',
    {x:0.25,y:1.18,w:9.5,h:0.28,
     fontSize:10,color:C_GRAY,fontFace:'Calibri',align:'center',italic:true});

  if (data.rel.length > 0) {
    const labs = data.rel.map(d=>d.label);
    sl.addChart(pres.charts.BAR, [
      {name:'Relaciones positivas IRT 1', labels:labs, values:data.rel.map(d=>d.pos_irt1)},
      {name:'Relaciones positivas IRT 2', labels:labs, values:data.rel.map(d=>d.pos_irt2)},
    ], {
      x:0.5,y:1.5,w:9.0,h:3.8,barDir:'col',barGrouping:'clustered',
      chartColors:[C_IRT1,C_IRT2],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:12,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:100,valAxisNumFmt:'0"%"',
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
      showLegend:true,legendPos:'b',legendFontSize:11,
    });
  }
  sl.addText(`N = ${data.meta.N_irt2}  ·  ${data.meta.servicio}  ·  ${data.meta.periodo}`,
    {x:0.25,y:5.35,w:9.5,h:0.25,
     fontSize:8.5,color:'AAAAAA',fontFace:'Calibri',align:'center',italic:true});
}

pres.writeFile({fileName: OUTPUT})
  .then(() => { console.log('✅  PowerPoint guardado: ' + OUTPUT); })
  .catch(e  => { console.error('Error JS:', e); process.exit(1); });
"""

js_path = '/home/claude/_irt_builder.js'
with open(js_path, 'w', encoding='utf-8') as f:
    f.write(JS_CODE)

print('\n→ Construyendo PowerPoint con Node.js + pptxgenjs...')
result = subprocess.run(['node', js_path], capture_output=True, text=True)
if result.returncode != 0:
    print('ERROR en Node.js:')
    print(result.stderr)
    sys.exit(1)
print(result.stdout.strip())

os.remove(json_path)
os.remove(js_path)

print('\n' + '='*60)
print(f'  ✅  LISTO  →  {OUTPUT_FILE}')
print(f'      {N_irt2} pacientes con IRT2  ·  {PERIODO}')
print('='*60)
