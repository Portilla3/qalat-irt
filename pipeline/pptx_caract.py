#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   SCRIPT_IRT_Universal_PPTX_Caracterizacion.py  —  v1.1                   ║
║   Genera presentación PowerPoint de caracterización al ingreso (IRT1)    ║
║   8 slides · Compatible con cualquier país IRT                            ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  CÓMO USAR:                                                                 ║
║  1. Sube este script + la base Wide IRT                                    ║
║  2. Escribe: "Ejecuta el PPTX Caracterización IRT"                        ║
║                                                                             ║
║  SLIDES:                                                                    ║
║    1. Portada                                                               ║
║    2. Antecedentes generales (KPIs + sexo + edad)                         ║
║    3. Sustancia principal (torta)                                          ║
║    4. Días consumo sustancia principal                                     ║
║    5. % Consumidores + Días promedio por sustancia                        ║
║    6. Urgencias + Accidentes + Salud                                      ║
║    7. Transgresión total + por tipo                                        ║
║    8. Relaciones interpersonales + Satisfacción de vida                   ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import glob, os, unicodedata

def _norm(s):
    return unicodedata.normalize('NFD', str(s).lower()).encode('ascii','ignore').decode()

# ── Detección de país ─────────────────────────────────────────────────────────
_PAISES = {
    'republica_dominicana':'República Dominicana','repdomini':'República Dominicana',
    'dominicana':'República Dominicana','honduras':'Honduras',
    'panama':'Panamá','panam':'Panamá','el_salvador':'El Salvador',
    'salvador':'El Salvador','mexico':'México','mexic':'México',
    'ecuador':'Ecuador','peru':'Perú','argentina':'Argentina',
    'colombia':'Colombia','chile':'Chile','bolivia':'Bolivia',
    'paraguay':'Paraguay','uruguay':'Uruguay','venezuela':'Venezuela',
    'guatemala':'Guatemala','costa_rica':'Costa Rica',
    'costarica':'Costa Rica','nicaragua':'Nicaragua',
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
print('  SCRIPT_IRT_Universal_PPTX_Caracterizacion  v1.1')
print('=' * 60)

INPUT_FILE  = auto_archivo_wide()
OUTPUT_FILE = '/home/claude/IRT_Presentacion_Caracterizacion.pptx'

# ── FILTRO OPCIONAL POR CENTRO ────────────────────────────────────────────────
FILTRO_CENTRO = None
# ─────────────────────────────────────────────────────────────────────────────

import pandas as pd, numpy as np, json, subprocess, sys, warnings
warnings.filterwarnings('ignore')

df = pd.read_excel(INPUT_FILE, sheet_name='Base Wide', header=1)
df.columns = [str(c) for c in df.columns]
cols = df.columns.tolist()

# Filtro de centro
_col_centro = next((c for c in cols if any(x in _norm(c) for x in
                    ['codigo del centro', 'servicio de tratamiento',
                     'centro/ servicio', 'codigo centro'])), None)
if FILTRO_CENTRO and _col_centro:
    n_antes = len(df)
    df = df[df[_col_centro].astype(str).str.strip() == FILTRO_CENTRO].copy()
    df = df.reset_index(drop=True)
    print(f'\n⚑  Filtro: "{FILTRO_CENTRO}"  ({n_antes} → {len(df)} pacientes)')
    OUTPUT_FILE = f'/home/claude/IRT_Presentacion_Caracterizacion_{FILTRO_CENTRO}.pptx'

# Solo IRT1
mask1 = df['Tiene_IRT1'] == 'Sí' if 'Tiene_IRT1' in cols else pd.Series([True]*len(df))
df1   = df[mask1].copy().reset_index(drop=True)
N     = len(df1)

print(f'\n→ {N} pacientes IRT1')

# País / servicio / período
PAIS = _detectar_pais(INPUT_FILE)
if FILTRO_CENTRO:
    SERVICIO = f'{PAIS}  —  Centro {FILTRO_CENTRO}' if PAIS else f'Centro {FILTRO_CENTRO}'
else:
    SERVICIO = PAIS if PAIS else 'Servicio de Tratamiento'

MESES = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
         7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'}
hoy = pd.Timestamp.now()
fecha_col = next((c for c in cols if 'fecha de administracion' in _norm(c)), None)
PERIODO = '2025'
if fecha_col:
    fch = pd.to_datetime(df1[fecha_col], errors='coerce').dropna()
    fch = fch[(fch.dt.year >= 2010) & (fch.dt.year <= hoy.year+1)]
    if len(fch):
        f0, f1_ = fch.min(), fch.max()
        PERIODO = (f'{MESES[f0.month]} {f0.year}'
                   if f0.year==f1_.year and f0.month==f1_.month
                   else f'{MESES[f0.month]}–{MESES[f1_.month]} {f0.year}'
                   if f0.year==f1_.year
                   else f'{MESES[f0.month]} {f0.year} – {MESES[f1_.month]} {f1_.year}')

print(f'  Servicio: {SERVICIO} | Período: {PERIODO}')

# ══════════════════════════════════════════════════════════════════════════════
# DETECCIÓN DE COLUMNAS (_IRT1)
# ══════════════════════════════════════════════════════════════════════════════
def col1(kws):
    for c in cols:
        if not c.endswith('_IRT1'): continue
        if all(_norm(k) in _norm(c) for k in kws): return c
    return None

COL_SEXO = next((c for c in cols if _norm(c) in ['sexo','género','genero']), None)
COL_FN   = next((c for c in cols if 'fecha de nacimiento' in _norm(c)), None)
COL_SP   = col1(['sustancia','principal'])

SUST_NOMBRES = {
    'Alcohol':['alcohol'],'Marihuana':['marihuana','cannabis'],
    'Heroína':['heroina'],'Cocaína':['cocain'],
    'Fentanilo':['fentanil'],'Inhalables':['inhalab'],
    'Metanfetamina':['metanfet','cristal'],'Crack':['crack'],
    'Pasta Base':['pasta base','pasta'],'Sedantes':['sedant','benzod'],
    'Opiáceos':['opiod','opiac'],'Tabaco':['tabaco','nicot'],
    'Otra sustancia':['otra sust'],
}
SUST_TOTAL1 = {}
for sust, kws in SUST_NOMBRES.items():
    for c in cols:
        if not c.endswith('_IRT1'): continue
        nc = _norm(c)
        if any(_norm(k) in nc for k in kws) and ('total' in nc or '(0-28)' in nc):
            SUST_TOTAL1[sust] = c; break
SUST_ACTIVAS = list(SUST_TOTAL1.keys())

COL_SPSI = col1(['salud','psicol'])
COL_SFIS = col1(['salud','fis'])
COL_URG  = next((c for c in cols if c.endswith('_IRT1') and '5)' in c and
                  any(k in c.lower() for k in ['urgencia','hospitali','emergencia'])), None)
COL_ACC  = next((c for c in cols if c.endswith('_IRT1') and '6)' in c and
                  'accidente' in c.lower()), None)

TRANS_DEF = {'Robo / Hurto':'robo','Venta de sustancias':'venta',
             'Violencia a otras personas':'violencia',
             'Violencia intrafamiliar':'intraf','Detenido / Arrestado':'detenido'}
TRANS_COLS = {n: next((c for c in cols if c.endswith('_IRT1') and kw in c.lower()), None)
              for n,kw in TRANS_DEF.items()}

REL_MAP = {'Padre':['padre'],'Madre':['madre'],'Hijos':['hijos','hijo'],
           'Hermanos':['hermanos'],'Pareja':['pareja'],'Amigos':['amigos'],'Otros':['otros']}
REL_COLS = {}
for vin, kws in REL_MAP.items():
    for c in cols:
        if not c.endswith('_IRT1'): continue
        nc = _norm(c)
        if '14)' not in c and 'relaci' not in nc: continue
        if any(_norm(k) in nc for k in kws): REL_COLS[vin]=c; break

SAT_MAP = {
    'Vida en general':[['16)'],['satisfac','vida']],
    'Lugar donde vive':[['17)'],['satisfac','lugar']],
    'Situación laboral':[['18)'],['satisfac','labor','educac']],
    'Tiempo libre':[['19)'],['satisfac','tiempo']],
    'Cap. económica':[['20)'],['satisfac','econom']],
}
SAT_COLS = {}
for dim,(nums,kws) in SAT_MAP.items():
    for c in cols:
        if not c.endswith('_IRT1'): continue
        nc = _norm(c)
        if any(n in c for n in nums) and any(_norm(k) in nc for k in kws):
            SAT_COLS[dim]=c; break

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

def prom1(col):
    if col is None: return None, 0
    v = pd.to_numeric(df1[col], errors='coerce').dropna()
    return (round(float(v.mean()),1) if len(v) else None), len(v)

def sino1(col):
    if col is None: return None, None
    v = df1[col].dropna().astype(str).str.strip().str.lower()
    nv = len(v); nsi = int(v.isin(['sí','si']).sum())
    return nsi, round(nsi/nv*100,1) if nv else 0.0

# Sexo
R_sexo = {}
if COL_SEXO:
    R_sexo = {k:int(v) for k,v in df1[COL_SEXO].value_counts(dropna=True).items()}
n_hombre = R_sexo.get('Hombre', R_sexo.get('H', max(R_sexo.values()) if R_sexo else 0))
n_mujer  = R_sexo.get('Mujer',  R_sexo.get('M', min(R_sexo.values()) if len(R_sexo)>1 else 0))
N_sx     = n_hombre + n_mujer
pct_h    = round(n_hombre/N_sx*100,1) if N_sx else 0

# Edad
edad_media = 0; edad_grupos = []
if COL_FN:
    edades = ((hoy - pd.to_datetime(df1[COL_FN], errors='coerce')).dt.days/365.25).dropna()
    edades = edades[(edades>=10)&(edades<=100)]
    if len(edades):
        edad_media = round(float(edades.mean()),1)
        bins = [0,17,25,35,45,55,200]
        labs = ['<18','18–25','26–35','36–45','46–55','56+']
        ec = pd.cut(edades,bins=bins,labels=labs)
        total_e = len(edades)
        edad_grupos = [{'label':l,'pct':round(int((ec==l).sum())/total_e*100,1)}
                       for l in labs if int((ec==l).sum())>0]

# Sustancia principal
sust_pp = df1[COL_SP].apply(norm_sust).dropna() if COL_SP else pd.Series(dtype=str)
R_sp = sust_pp.value_counts().to_dict() if len(sust_pp) else {}
N_sp = len(sust_pp)
sust_top1     = sust_pp.value_counts().index[0] if N_sp else '—'
sust_top1_pct = round(sust_pp.value_counts().iloc[0]/N_sp*100,1) if N_sp else 0

# Días consumo por sustancia PRINCIPAL (solo quienes la declaran como principal)
dias_pp = []
for sust, col in SUST_TOTAL1.items():
    msk = sust_pp == sust
    vals = pd.to_numeric(df1.loc[msk.reindex(df1.index,fill_value=False), col],
                         errors='coerce').dropna()
    vals = vals[vals>0]
    if len(vals): dias_pp.append({'label':sust,'prom':round(float(vals.mean()),1),'n':int(len(vals))})
dias_pp.sort(key=lambda x:-x['prom'])

# % consumidores
cons_pct = []
for sust, col in SUST_TOTAL1.items():
    v = pd.to_numeric(df1[col], errors='coerce').fillna(0)
    n_c = int((v>0).sum())
    if n_c: cons_pct.append({'label':sust,'pct':round(n_c/N*100,1),'n':n_c})
cons_pct.sort(key=lambda x:-x['pct'])

# Días promedio por sustancia (todos los consumidores)
dias_sust = []
for sust, col in SUST_TOTAL1.items():
    v = pd.to_numeric(df1[col], errors='coerce'); sub = v[v>0].dropna()
    if len(sub): dias_sust.append({'label':sust,'prom':round(float(sub.mean()),1),'n':int(len(sub))})
dias_sust.sort(key=lambda x:-x['prom'])

# Urgencias y accidentes
n_urg, pct_urg = sino1(COL_URG)
n_acc, pct_acc = sino1(COL_ACC)

# Salud
R_salud = []
for nombre, col in [('Salud Psicológica (0–10)', COL_SPSI),
                    ('Salud Física (0–10)',       COL_SFIS)]:
    m, nv = prom1(col)
    if m is not None: R_salud.append({'label':nombre,'prom':m})

# Transgresión
# Detectar formato: Sí/No dicotómico  vs  numérico (nº de veces)
def _trans_positivo(series):
    """Devuelve Serie booleana: True si el paciente cometió la transgresión.
       Soporta formato Sí/No Y formato numérico (valor > 0)."""
    s = series.dropna()
    muestra = s.astype(str).str.lower().head(20)
    es_dicotomica = muestra.isin(['sí','si','no','no aplica']).any()
    if es_dicotomica:
        return series.astype(str).str.lower().isin(['sí','si'])
    else:
        # Formato numérico: reemplazar texto ('nunca','no','no aplica') por 0
        num = pd.to_numeric(
            series.astype(str).str.lower()
                  .str.replace('nunca','0').str.replace('no aplica','0')
                  .str.replace('no','0').str.strip(),
            errors='coerce').fillna(0)
        return num > 0

any_tr_series = []
for col in [c for c in TRANS_COLS.values() if c]:
    any_tr_series.append(_trans_positivo(df1[col]))
if any_tr_series:
    any_tr = pd.concat(any_tr_series, axis=1).any(axis=1)
else:
    any_tr = pd.Series(False, index=df1.index)
n_tr = int(any_tr.sum()); pct_tr = round(n_tr/N*100,1)

trans_tipos = []
for nombre, col in TRANS_COLS.items():
    if col is None: continue
    mask = _trans_positivo(df1[col])
    nsi = int(mask.sum())
    if nsi: trans_tipos.append({'label':nombre,'n':nsi,'pct':round(nsi/N*100,1)})

# Relaciones interpersonales (% positivas)
R_rel = []
for vin, col in REL_COLS.items():
    vals = df1[col].dropna()
    vals = vals[vals.astype(str).str.lower() != 'no aplica']
    nv = len(vals)
    if nv == 0: continue
    pos = round(int(vals.astype(str).isin(['Excelente','Buena']).sum())/nv*100,1)
    R_rel.append({'label':vin,'pos':pos,'n':nv})

# Satisfacción de vida
R_sat = []
for dim, col in SAT_COLS.items():
    m, nv = prom1(col)
    if m is not None: R_sat.append({'label':dim,'prom':m})

print(f'  ✓ Calculados | Sust: {SUST_ACTIVAS}')
print(f'  Trans: {pct_tr}% | Urgencias: {pct_urg}% | Accidentes: {pct_acc}%')

# ══════════════════════════════════════════════════════════════════════════════
# JSON
# ══════════════════════════════════════════════════════════════════════════════
data = {
    'meta': {
        'servicio':       SERVICIO,
        'periodo':        PERIODO,
        'N':              N,
        'pct_h':          pct_h,
        'n_h':            n_hombre,
        'edad_media':     edad_media,
        'sust_top1':      sust_top1,
        'sust_top1_pct':  sust_top1_pct,
    },
    'sexo':      [{'label':k,'n':v} for k,v in R_sexo.items()],
    'edad':      edad_grupos,
    'sust':      [{'label':k,'pct':round(v/N_sp*100,1),'n':v} for k,v in R_sp.items()],
    'dias_pp':   dias_pp,
    'consumo':   cons_pct,
    'dias':      dias_sust,
    'urgencias': {'n': n_urg or 0, 'pct': pct_urg or 0},
    'accidentes':{'n': n_acc or 0, 'pct': pct_acc or 0},
    'salud':     R_salud,
    'transTotal':{'n': n_tr, 'pct': pct_tr},
    'transTipos':trans_tipos,
    'rel':       R_rel,
    'sat':       R_sat,
}

json_path = '/home/claude/_irt_car_data.json'
with open(json_path,'w',encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

# ══════════════════════════════════════════════════════════════════════════════
# NODE.JS — pptxgenjs
# ══════════════════════════════════════════════════════════════════════════════
JS_CODE = r"""
const fs      = require('fs');
const pptxgen = require('pptxgenjs');

const data   = JSON.parse(fs.readFileSync('/home/claude/_irt_car_data.json', 'utf8'));
const OUTPUT = '""" + OUTPUT_FILE + r"""';

const C_DARK  = '1F3864', C_MID = '2E75B6', C_LIGHT = 'BDD7EE';
const C_IRT1  = '1F3864';
const C_TITLE = '0070C0', C_GRAY = '595959', C_WHITE = 'FFFFFF';
const PIE_COLS= ['2E75B6','1F3864','00B0F0','9DC3E6','70AD47','4472C4',
                 'D9D9D9','C00000','ED7D31','FFC000','7030A0','538135'];

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

function hdr(sl, txt) {
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:10,h:0.72,
    fill:{color:C_DARK},line:{color:C_DARK}});
  sl.addShape(pres.shapes.RECTANGLE, {x:5.5,y:0,w:4.5,h:0.72,
    fill:{color:C_MID,transparency:40},line:{color:C_MID,transparency:40}});
  sl.addText(txt, {x:0.25,y:0,w:9.5,h:0.72,
    fontSize:22,bold:true,color:C_WHITE,fontFace:'Calibri',valign:'middle'});
}
function divV(sl, x) {
  sl.addShape(pres.shapes.LINE, {x,y:0.78,w:0,h:4.85,
    line:{color:'D9D9D9',width:1}});
}

const TITULO = `Caracterización al Ingreso · ${data.meta.servicio}`;

// ── SLIDE 1: PORTADA ──────────────────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  sl.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:4.0,h:5.625,
    fill:{color:C_DARK},line:{color:C_DARK}});
  sl.addShape(pres.shapes.RECTANGLE, {x:3.1,y:0,w:1.5,h:5.625,
    fill:{color:C_MID,transparency:60},line:{color:C_MID,transparency:60}});
  sl.addText('Caracterización', {x:0.25,y:1.6,w:3.2,h:0.7,
    fontSize:22,bold:true,color:C_WHITE,fontFace:'Calibri'});
  sl.addText('Ingreso a Tratamiento · IRT', {x:0.25,y:2.35,w:3.2,h:0.55,
    fontSize:12,color:C_LIGHT,fontFace:'Calibri'});
  sl.addText([
    {text:'IRT 1', options:{breakLine:true}},
    {text:'Ingreso a Tratamiento'}
  ], {x:4.6,y:1.55,w:5.1,h:1.4,
    fontSize:30,bold:true,color:C_GRAY,fontFace:'Calibri',align:'center',valign:'middle'});
  sl.addText(data.meta.servicio.toUpperCase(), {x:4.6,y:3.1,w:5.1,h:0.45,
    fontSize:17,bold:true,color:C_MID,fontFace:'Calibri',align:'center'});
  sl.addText(data.meta.periodo, {x:4.6,y:3.62,w:5.1,h:0.35,
    fontSize:13,bold:true,color:C_MID,fontFace:'Calibri',align:'center'});
  sl.addText(`N = ${data.meta.N} personas al ingreso a tratamiento`,
    {x:4.6,y:4.1,w:5.1,h:0.35,fontSize:11,color:'888888',fontFace:'Calibri',align:'center'});
}

// ── SLIDE 2: ANTECEDENTES GENERALES ──────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  divV(sl, 4.95);

  // KPIs: N · % hombres · edad promedio
  const kpis = [
    {val: String(data.meta.N),          lab:'Personas\ningresaron'},
    {val: `${data.meta.pct_h}%`,         lab:'Son\nhombres'},
    {val: String(data.meta.edad_media),  lab:'Edad\npromedio'},
  ];
  kpis.forEach((k,i) => {
    const x = 0.18 + i*1.55;
    sl.addShape(pres.shapes.RECTANGLE, {x,y:0.82,w:1.42,h:0.88,
      fill:{color:'EEF4FB'},line:{color:'BDD7EE',width:0.5}});
    sl.addText(k.val, {x,y:0.86,w:1.42,h:0.48,
      fontSize:20,bold:true,color:C_DARK,fontFace:'Calibri',align:'center',valign:'middle'});
    sl.addText(k.lab, {x,y:1.34,w:1.42,h:0.34,
      fontSize:9,color:C_GRAY,fontFace:'Calibri',align:'center',valign:'top'});
  });

  // Izquierda: torta sexo
  sl.addText('Distribución por Sexo',
    {x:0.25,y:1.85,w:4.5,h:0.35,
     fontSize:12,bold:true,color:C_GRAY,fontFace:'Calibri'});
  const sexoFilt = data.sexo.filter(s=>s.n>0);
  if (sexoFilt.length > 0) {
    sl.addChart(pres.charts.PIE, [{
      name:'Sexo',
      labels: sexoFilt.map(s=>s.label),
      values: sexoFilt.map(s=>s.n),
    }], {
      x:0.4,y:2.22,w:4.2,h:3.15,
      showPercent:true,showLabel:false,showLegend:true,legendPos:'b',legendFontSize:11,
      dataLabelFontSize:13,chartColors:['2E75B6','9DC3E6'],
      chartArea:{fill:{color:'FFFFFF'}},dataLabelColor:C_WHITE,
    });
  }

  // Derecha: edad
  sl.addText('Distribución por Rango de Edad',
    {x:5.15,y:0.85,w:4.6,h:0.35,
     fontSize:12,bold:true,color:C_GRAY,fontFace:'Calibri'});
  if (data.meta.edad_media > 0) {
    sl.addText(`Promedio: ${data.meta.edad_media} años`,
      {x:5.15,y:1.22,w:4.6,h:0.28,
       fontSize:10,color:C_GRAY,fontFace:'Calibri',italic:true});
  }
  if (data.edad.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'Personas',
      labels: data.edad.map(e=>e.label),
      values: data.edad.map(e=>e.pct),
    }], {
      x:5.1,y:1.52,w:4.65,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_IRT1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:10,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:11,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisNumFmt:'0"%"',
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }
}

// ── SLIDE 3: SUSTANCIA PRINCIPAL ─────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addText('CONSUMO SUSTANCIA PRINCIPAL AL INGRESO',
    {x:1.5,y:0.82,w:7,h:0.38,
     fontSize:14,bold:true,color:C_TITLE,fontFace:'Calibri',align:'center'});
  if (data.sust.length > 0) {
    sl.addChart(pres.charts.PIE, [{
      name:'Sustancia',
      labels: data.sust.map(s=>s.label),
      values: data.sust.map(s=>s.pct),
    }], {
      x:1.3,y:1.28,w:7.4,h:4.1,
      showPercent:true,showLabel:false,showLegend:true,legendPos:'b',legendFontSize:10,
      dataLabelFontSize:11,chartColors:PIE_COLS.slice(0,data.sust.length),
      chartArea:{fill:{color:'FFFFFF'}},dataLabelColor:C_WHITE,dataLabelPosition:'bestFit',
    });
    sl.addText(
      `Sustancia más frecuente: ${data.meta.sust_top1} (${data.meta.sust_top1_pct}%)  ·  N = ${data.meta.N}`,
      {x:1.0,y:5.3,w:8,h:0.25,
       fontSize:9,color:C_GRAY,fontFace:'Calibri',align:'center',italic:true});
  }
}

// ── SLIDE 4: DÍAS CONSUMO SUSTANCIA PRINCIPAL ─────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  sl.addText('PROMEDIO DE DÍAS DE CONSUMO DE LA SUSTANCIA PRINCIPAL\nÚltimas 4 semanas · solo personas cuya sust. principal corresponde a cada categoría',
    {x:0.5,y:0.82,w:9.0,h:0.65,
     fontSize:12,bold:true,color:C_TITLE,fontFace:'Calibri',align:'center'});
  if (data.dias_pp.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'Días promedio',
      labels: data.dias_pp.map(d=>d.label),
      values: data.dias_pp.map(d=>d.prom),
    }], {
      x:1.2,y:1.55,w:7.6,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_IRT1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:11,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:28,valAxisMinVal:0,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }
  sl.addText(`N = ${data.meta.N}  ·  Escala: días en últimas 4 semanas (0–28)`,
    {x:0.25,y:5.35,w:9.5,h:0.25,
     fontSize:8.5,color:'AAAAAA',fontFace:'Calibri',align:'center',italic:true});
}

// ── SLIDE 5: % CONSUMIDORES + DÍAS POR SUSTANCIA ──────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  divV(sl, 4.95);

  sl.addText('% DE PERSONAS QUE CONSUME\nCada sustancia al ingreso',
    {x:0.25,y:0.82,w:4.5,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  if (data.consumo.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'% Consumidores',
      labels: data.consumo.map(d=>d.label),
      values: data.consumo.map(d=>d.pct),
    }], {
      x:0.2,y:1.52,w:4.5,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_IRT1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:10,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:9,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:100,valAxisNumFmt:'0"%"',
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }

  sl.addText('PROMEDIO DE DÍAS DE CONSUMO\nPor sustancia (solo consumidores)',
    {x:5.15,y:0.82,w:4.6,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  if (data.dias.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'Días promedio',
      labels: data.dias.map(d=>d.label),
      values: data.dias.map(d=>d.prom),
    }], {
      x:5.15,y:1.52,w:4.6,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_IRT1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFontSize:10,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:9,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:28,valAxisMinVal:0,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }
  sl.addText(`N = ${data.meta.N}  ·  Una persona puede consumir más de una sustancia`,
    {x:0.25,y:5.35,w:9.5,h:0.25,
     fontSize:8.5,color:'AAAAAA',fontFace:'Calibri',align:'center',italic:true});
}

// ── SLIDE 6: URGENCIAS + ACCIDENTES + SALUD ───────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  divV(sl, 4.95);

  const U = data.urgencias; const A = data.accidentes;

  // Izquierda: urgencias + accidentes (2 barras juntas)
  sl.addText('Urgencias / Hospitalizaciones\ny Accidentes por consumo',
    {x:0.25,y:0.82,w:4.5,h:0.65,
     fontSize:12,bold:true,color:C_GRAY,fontFace:'Calibri',align:'left'});
  const uaLabels = ['Urgencias /\nHospitalización','Accidentes\npor consumo'];
  const uaVals   = [U.pct, A.pct];
  sl.addChart(pres.charts.BAR, [{
    name:'% personas',
    labels: uaLabels,
    values: uaVals,
  }], {
    x:0.3,y:1.52,w:4.4,h:3.85,barDir:'col',barGrouping:'clustered',
    chartColors:[C_IRT1],chartArea:{fill:{color:'FFFFFF'}},
    showValue:true,dataLabelFormatCode:'0.0"%"',dataLabelFontSize:14,dataLabelColor:'363636',
    catAxisLabelColor:'363636',catAxisLabelFontSize:11,
    valAxisLabelColor:'595959',valAxisLabelFontSize:9,
    valAxisMaxVal:Math.max(U.pct,A.pct,5)*1.45,
    valAxisNumFmt:'0"%"',
    valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
  });

  // Derecha: salud
  sl.addText('AUTOPERCEPCIÓN DEL ESTADO DE SALUD\n(escala 0–10)',
    {x:5.15,y:0.82,w:4.6,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  if (data.salud.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'Promedio',
      labels: data.salud.map(d=>d.label),
      values: data.salud.map(d=>d.prom),
    }], {
      x:5.15,y:1.52,w:4.6,h:3.85,barDir:'bar',barGrouping:'clustered',
      chartColors:[C_IRT1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:10,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,valAxisMaxVal:10,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }
  sl.addText(`N = ${data.meta.N}`,
    {x:0.25,y:5.35,w:9.5,h:0.25,
     fontSize:8.5,color:'AAAAAA',fontFace:'Calibri',align:'center',italic:true});
}

// ── SLIDE 7: TRANSGRESIÓN ─────────────────────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  divV(sl, 4.95);
  sl.addText('Personas que cometieron alguna\ntransgresión a la norma social',
    {x:0.25,y:0.82,w:4.5,h:0.65,
     fontSize:13,bold:true,color:C_GRAY,fontFace:'Calibri',align:'left'});
  sl.addText('Distribución por tipo de transgresión',
    {x:5.15,y:0.82,w:4.6,h:0.65,
     fontSize:13,bold:true,color:C_GRAY,fontFace:'Calibri',align:'center'});

  const T = data.transTotal;
  sl.addChart(pres.charts.BAR, [
    {name:'Con transgresión', labels:['Con\ntransgresión','Sin\ntransgresión'],
     values:[T.pct, null]},
    {name:'Sin transgresión', labels:['Con\ntransgresión','Sin\ntransgresión'],
     values:[null, parseFloat((100-T.pct).toFixed(1))]},
  ], {
    x:0.3,y:1.52,w:4.4,h:3.85,barDir:'col',barGrouping:'clustered',
    chartColors:[C_DARK,'BDD7EE'],chartArea:{fill:{color:'FFFFFF'}},
    showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:14,
    dataLabelColor:C_WHITE,dataLabelPosition:'inEnd',
    catAxisLabelColor:'363636',catAxisLabelFontSize:11,
    valAxisLabelColor:'595959',valAxisLabelFontSize:9,
    valAxisMaxVal:100,valAxisNumFmt:'0"%"',
    valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
  });
  sl.addText(`${T.n} personas (${T.pct}%)`,
    {x:0.3,y:1.22,w:4.4,h:0.28,
     fontSize:12,bold:true,color:C_DARK,fontFace:'Calibri',align:'center'});

  const tiposFilt = data.transTipos.filter(d=>d.pct>0);
  if (tiposFilt.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'% personas',
      labels: tiposFilt.map(d=>d.label),
      values: tiposFilt.map(d=>d.pct),
    }], {
      x:5.1,y:1.52,w:4.65,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_IRT1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:9,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisNumFmt:'0"%"',
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }
  sl.addText(`N = ${data.meta.N}`,
    {x:0.25,y:5.35,w:9.5,h:0.25,
     fontSize:8.5,color:'AAAAAA',fontFace:'Calibri',align:'center',italic:true});
}

// ── SLIDE 8: RELACIONES + SATISFACCIÓN ───────────────────────────────────
{
  const sl = pres.addSlide(); sl.background = {color:'FFFFFF'};
  hdr(sl, TITULO);
  divV(sl, 4.95);

  // Izquierda: relaciones (% positivas)
  sl.addText('CALIDAD DE RELACIONES INTERPERSONALES\n% positivas (Excelente + Buena)',
    {x:0.25,y:0.82,w:4.5,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  if (data.rel.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'% Positivas',
      labels: data.rel.map(d=>d.label),
      values: data.rel.map(d=>d.pos),
    }], {
      x:0.2,y:1.52,w:4.5,h:3.85,barDir:'col',barGrouping:'clustered',
      chartColors:[C_IRT1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFormatCode:'0"%"',dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:10,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,
      valAxisMaxVal:100,valAxisNumFmt:'0"%"',
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }

  // Derecha: satisfacción
  sl.addText('SATISFACCIÓN DE VIDA\n(escala 0–10)',
    {x:5.15,y:0.82,w:4.6,h:0.65,
     fontSize:11,bold:true,color:C_TITLE,fontFace:'Calibri',align:'left'});
  if (data.sat.length > 0) {
    sl.addChart(pres.charts.BAR, [{
      name:'Promedio',
      labels: data.sat.map(d=>d.label),
      values: data.sat.map(d=>d.prom),
    }], {
      x:5.15,y:1.52,w:4.6,h:3.85,barDir:'bar',barGrouping:'clustered',
      chartColors:[C_IRT1],chartArea:{fill:{color:'FFFFFF'}},
      showValue:true,dataLabelFontSize:11,dataLabelColor:'363636',
      catAxisLabelColor:'363636',catAxisLabelFontSize:10,
      valAxisLabelColor:'595959',valAxisLabelFontSize:9,valAxisMaxVal:10,
      valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},showLegend:false,
    });
  }
  sl.addText(`N = ${data.meta.N}  ·  ${data.meta.servicio}  ·  ${data.meta.periodo}`,
    {x:0.25,y:5.35,w:9.5,h:0.25,
     fontSize:8.5,color:'AAAAAA',fontFace:'Calibri',align:'center',italic:true});
}

pres.writeFile({fileName: OUTPUT})
  .then(() => { console.log('✅  PowerPoint guardado: ' + OUTPUT); })
  .catch(e  => { console.error('Error JS:', e); process.exit(1); });
"""

js_path = '/home/claude/_irt_car_builder.js'
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
print(f'      {N} pacientes IRT1  ·  {PERIODO}')
print('='*60)
