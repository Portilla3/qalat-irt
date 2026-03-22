"""
app.py — QALAT · Sistema de Monitoreo IRT
v1.0 — Instrumento IRT
"""
import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import tempfile, os, sys
from io import BytesIO
from datetime import datetime, date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pipeline.wide_irt    import procesar_wide
from pipeline.runner_irt  import run_script, run_paquetes_centros

NAVY='#1F3864'; MID='#2E75B6'; ACCENT='#00B0F0'
ORANGE='#C8590A'; RED='#C00000'; GREEN='#538135'

st.set_page_config(page_title='QALAT · IRT', page_icon='📊',
                   layout='wide', initial_sidebar_state='collapsed')
st.markdown(f"""<style>
html,body,[class*="css"]{{font-family:'Calibri',sans-serif;}}
.main{{background:#F8FAFD;}}
.qalat-hdr{{background:{NAVY};color:white;padding:1.2rem 2rem;border-radius:8px;margin-bottom:1.5rem;}}
.qalat-hdr h1{{color:white;font-size:1.6rem;margin:0;}}
.qalat-hdr p{{color:#BDD7EE;font-size:.9rem;margin:.3rem 0 0 0;}}
.kpi{{background:white;border-radius:8px;padding:1rem 1.2rem;border-left:4px solid {MID};
      box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:.5rem;}}
.kpi.red{{border-left-color:{RED};}}.kpi.orange{{border-left-color:{ORANGE};}}.kpi.green{{border-left-color:{GREEN};}}
.kpi-lbl{{font-size:.78rem;color:#666;margin-bottom:.2rem;}}
.kpi-val{{font-size:1.8rem;font-weight:700;color:{NAVY};}}
.kpi-sub{{font-size:.75rem;color:#888;}}
.sec{{background:{MID};color:white;padding:.5rem 1rem;border-radius:6px;
      font-weight:600;font-size:1rem;margin:1.2rem 0 .8rem 0;}}
.filter-box{{background:white;border:1px solid #D0DFF0;border-radius:8px;padding:1rem 1.2rem;margin-bottom:1rem;}}
.filter-box h4{{color:{NAVY};margin:0 0 .6rem 0;font-size:.95rem;}}
.outcard{{background:white;border-radius:8px;padding:1rem;border:1px solid #D0DFF0;margin-bottom:.5rem;}}
.outcard h4{{color:{NAVY};margin:0 0 .3rem 0;font-size:.95rem;}}
.outcard p{{color:#666;font-size:.8rem;margin:0;}}
div.stButton>button{{background:#1E7E34;color:white;border:none;
    padding:.6rem 2rem;border-radius:6px;font-size:1rem;font-weight:600;width:100%;
    box-shadow:0 2px 6px rgba(30,126,52,.35);letter-spacing:.3px;}}
div.stButton>button:hover{{background:#145222;box-shadow:0 3px 10px rgba(30,126,52,.5);}}
#MainMenu,footer,header{{visibility:hidden;}}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="qalat-hdr">
  <h1>📊 QALAT · Monitoreo de Resultados de Tratamiento — Instrumento IRT</h1>
  <p>Procesamiento automático IRT · Sube tu Excel, aplica filtros y descarga todos los reportes</p>
  <p style="margin-top:.6rem;font-size:.75rem;color:#7fa8cc;">© Rodrigo Portilla · UNODC</p>
</div>""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown('### 📋 Pasos')
    st.markdown('1. Sube tu Excel bruto\n2. Aplica filtros (opcional)\n3. Elige reportes\n4. Clic en **Procesar**\n5. Descarga')
    st.markdown('---')
    st.caption(f'QALAT IRT v1.0 · {datetime.now().strftime("%d/%m/%Y")}')
    st.markdown('---')
    st.markdown(
        '<div style="font-size:.75rem;color:#999;line-height:1.6;">'
        '© Rodrigo Portilla<br>'
        '<span style="color:#bbb;">UNODC Chile · Proyecto QALAT</span>'
        '</div>', unsafe_allow_html=True)

LABELS = {
    'caract_excel':('📋 Tablas caracterización', 'Excel',      'Tablas al ingreso: sexo, edad, sustancias, salud, transgresión'),
    'seg_excel':   ('📋 Tablas seguimiento',      'Excel',      'Comparativo IRT1 vs IRT2'),
    'word_caract': ('📄 Word caracterización',    'Word',       'Informe Word al ingreso'),
    'word_seg':    ('📄 Word seguimiento',        'Word',       'Comparativo ingreso vs seguimiento'),
}

# ── Carga ─────────────────────────────────────────────────────────────────────
st.markdown('<div class="sec">📁 Cargar base de datos</div>', unsafe_allow_html=True)
uploaded = st.file_uploader('Arrastra tu Excel IRT aquí o haz clic para buscar',
                             type=['xlsx','xls'],
                             help='Archivo bruto exportado de Jotform — instrumento IRT')

filtro_centro_val = None
fecha_desde_val   = None
fecha_hasta_val   = None
centros_disponibles = []

if uploaded:
    @st.cache_data(show_spinner=False)
    def _leer_preview(file_bytes):
        import pandas as _pd, io, unicodedata
        def _n(s): return unicodedata.normalize('NFD',str(s).lower()).encode('ascii','ignore').decode()
        df = _pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=0)
        df.columns = [str(c) for c in df.columns]
        col_c = next((c for c in df.columns if any(k in _n(c) for k in
                      ['codigo del centro','centro de tratamiento','servicio de tratamiento'])
                      and 'trabajo' not in _n(c) and 'estudio' not in _n(c)), None)
        col_f = next((c for c in df.columns if any(k in _n(c) for k in
                      ['fecha de administracion','fecha_administracion','fecha entrevista'])), None)
        centros = sorted(df[col_c].dropna().astype(str).str.strip().unique().tolist()) if col_c else []
        fechas  = _pd.to_datetime(df[col_f], errors='coerce').dropna() if col_f else _pd.Series([], dtype='datetime64[ns]')
        return centros, fechas

    file_bytes = uploaded.getvalue()
    centros_disponibles, fechas_serie = _leer_preview(file_bytes)

    st.markdown('<div class="sec">🔍 Filtros (opcional)</div>', unsafe_allow_html=True)
    fc1, fc2, fc3 = st.columns([1.5, 1.5, 1])

    with fc1:
        st.markdown('<div class="filter-box"><h4>🏥 Filtrar por centro</h4>', unsafe_allow_html=True)
        sel_centro = st.selectbox('Centro', ['Todos los centros'] + centros_disponibles,
                                  label_visibility='collapsed')
        if sel_centro != 'Todos los centros': filtro_centro_val = sel_centro
        st.markdown('</div>', unsafe_allow_html=True)

    with fc2:
        st.markdown('<div class="filter-box"><h4>📅 Filtrar por período</h4>', unsafe_allow_html=True)
        MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
        anio_actual = datetime.now().year
        if len(fechas_serie):
            anio_min = max(fechas_serie.dt.year.min(), anio_actual-10)
            anio_max = min(fechas_serie.dt.year.max(), anio_actual+1)
        else:
            anio_min, anio_max = anio_actual-3, anio_actual
        anios = list(range(int(anio_min), int(anio_max)+1))
        p1, p2 = st.columns(2)
        with p1:
            st.caption('Desde')
            mes_d  = st.selectbox('Mes inicio', MESES, index=0,   key='mes_d', label_visibility='collapsed')
            anio_d = st.selectbox('Año inicio', anios, index=0,   key='anio_d', label_visibility='collapsed')
        with p2:
            st.caption('Hasta')
            mes_h  = st.selectbox('Mes fin',   MESES, index=11,   key='mes_h', label_visibility='collapsed')
            anio_h = st.selectbox('Año fin',   anios, index=len(anios)-1, key='anio_h', label_visibility='collapsed')
        usar_periodo = st.checkbox('Aplicar filtro de período', value=False)
        if usar_periodo:
            fecha_desde_val = f'{anio_d}-{MESES.index(mes_d)+1:02d}'
            fecha_hasta_val = f'{anio_h}-{MESES.index(mes_h)+1:02d}'
        st.markdown('</div>', unsafe_allow_html=True)

    with fc3:
        st.markdown('<div class="filter-box"><h4>📄 Reportes a generar</h4>', unsafe_allow_html=True)
        cb_ce  = st.checkbox('Tablas caracterización', value=False, key='cb_ce')
        cb_se  = st.checkbox('Tablas seguimiento',     value=False, key='cb_se')
        cb_wc  = st.checkbox('Word caracterización',   value=False, key='cb_wc')
        cb_ws  = st.checkbox('Word seguimiento',       value=False, key='cb_ws')
        st.markdown('</div>', unsafe_allow_html=True)

    SELECCION = {
        'caract_excel': cb_ce, 'seg_excel': cb_se,
        'word_caract':  cb_wc, 'word_seg':  cb_ws,
    }

    badges = ''
    if filtro_centro_val: badges += f'<span style="background:#E8F0FE;color:{NAVY};padding:3px 10px;border-radius:12px;font-size:.78rem;font-weight:600;margin-right:4px;">🏥 Centro: {filtro_centro_val}</span>'
    if fecha_desde_val:   badges += f'<span style="background:#E8F5E9;color:#1B5E20;padding:3px 10px;border-radius:12px;font-size:.78rem;font-weight:600;">📅 {fecha_desde_val} → {fecha_hasta_val}</span>'
    if not badges: badges = '<span style="color:#888;font-size:.85rem">Sin filtros — procesa toda la base</span>'
    st.markdown(f'**Archivo:** `{uploaded.name}` &nbsp;|&nbsp; {badges}', unsafe_allow_html=True)

    # ── Botón procesar ─────────────────────────────────────────────────────────
    if st.button('⚡ Procesar y generar reportes', use_container_width=True):
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp.write(uploaded.read()); tmp_raw = tmp.name
        work_dir = tempfile.mkdtemp(prefix='qalat_irt_')

        try:
            with st.spinner('Paso 1 — Procesando base Wide IRT...'):
                result = procesar_wide(
                    tmp_raw,
                    filtro_centro=filtro_centro_val,
                    fecha_desde=fecha_desde_val,
                    fecha_hasta=fecha_hasta_val,
                )
                st.session_state['result']   = result
                st.session_state['filename'] = uploaded.name
                st.session_state['seleccion']= SELECCION

                wide_path = os.path.join(work_dir, 'IRT_Base_Wide.xlsx')
                with open(wide_path,'wb') as f:
                    f.write(result['excel_bytes'].getvalue())
                st.session_state['wide_path'] = wide_path
                st.session_state['work_dir']  = work_dir

            s = result['stats']
            st.success(f"✅ Base Wide IRT — {s['N_total']} pacientes · {s['N_irt2']} con IRT2 · {result['periodo']}")

            outputs = {}
            keys_sel = [k for k,v in SELECCION.items() if v]
            prog = st.progress(0, text='Generando reportes...')
            for i, key in enumerate(keys_sel):
                lbl = LABELS[key][0]
                prog.progress(i/len(keys_sel) if keys_sel else 1, text=f'Generando {lbl}...')
                try:
                    buf, fname, mime = run_script(key, wide_path, filtro_centro=filtro_centro_val)
                    outputs[key] = {'ok':True,'buf':buf,'fname':fname,'mime':mime}
                except Exception as e:
                    outputs[key] = {'ok':False,'error':str(e)}
            prog.progress(1.0, text='✅ Listo')
            st.session_state['outputs'] = outputs

        except Exception as e:
            st.error(f'❌ Error: {e}')
        finally:
            os.unlink(tmp_raw)

# ── Resultados ────────────────────────────────────────────────────────────────
if 'result' in st.session_state:
    R = st.session_state['result']
    s = R['stats']; wide = R['wide']
    fc = R.get('filtro_centro'); fd = R.get('fecha_desde'); fh = R.get('fecha_hasta')

    filtro_str = ''
    if fc: filtro_str += f' · Centro: {fc}'
    if fd: filtro_str += f' · {fd} → {fh}'

    st.markdown('---')
    st.markdown(f'<div class="sec">📊 Resultados IRT — {R["periodo"]}{filtro_str}</div>',
                unsafe_allow_html=True)

    # KPIs
    k1,k2,k3,k4,k5,k6 = st.columns(6)
    for col,lbl,val,sub,cls in [
        (k1,'Pacientes únicos',       s['N_total'], '',                           ''),
        (k2,'Con seguimiento IRT2',   s['N_irt2'],  f"{s['pct_irt2']}% del total",''),
        (k3,'Solo IRT1 (pendientes)', s['N_solo1'], '',                           ''),
        (k4,'Valores corregidos',     s['N_alertas'],'','red' if s['N_alertas'] else 'green'),
        (k5,'🔴 Urgentes (90+ días)', s['n_rojo'],  '', 'red'),
        (k6,'🟠 Próximos (60–89d)',   s['n_naranja'],'','orange'),
    ]:
        with col:
            st.markdown(f'<div class="kpi {cls}"><div class="kpi-lbl">{lbl}</div>'
                        f'<div class="kpi-val">{val}</div>'
                        f'{"<div class=kpi-sub>"+sub+"</div>" if sub else ""}</div>',
                        unsafe_allow_html=True)

    # Tabla centros
    centros = R.get('centros',[])
    if centros and not fc:
        st.markdown('<div class="sec">🏥 Resumen por Centro</div>', unsafe_allow_html=True)
        df_c = pd.DataFrame(centros)
        df_c.columns = ['Centro','Aplicaciones','Pacientes únicos','Con IRT2','Sin IRT2','Valores corregidos']
        rows_html = ''
        for i, row in df_c.iterrows():
            is_total = str(row.iloc[0]) == 'TOTAL'
            bg = f'background:{NAVY};color:white;font-weight:700;' if is_total else \
                 ('background:#EEF4FB;' if i%2==0 else 'background:white;')
            cells = ''
            for j, val in enumerate(row):
                align = 'left' if j==0 else 'center'
                cells += f'<td style="padding:7px 12px;text-align:{align};">{val}</td>'
            rows_html += f'<tr style="{bg}">{cells}</tr>'
        hdrs = ''.join(f'<th style="padding:9px 12px;text-align:{"left" if i==0 else "center"};background:{NAVY};color:white;font-size:.85rem;">{c}</th>'
                       for i,c in enumerate(df_c.columns))
        st.markdown(f'<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-family:Calibri,sans-serif;font-size:.9rem;"><thead><tr>{hdrs}</tr></thead><tbody>{rows_html}</tbody></table></div>',
                    unsafe_allow_html=True)

    # Gráficos
    st.markdown('<div class="sec">📈 Análisis visual</div>', unsafe_allow_html=True)
    gc1,gc2,gc3 = st.columns(3)

    with gc1:
        fig,ax=plt.subplots(figsize=(4.5,3.2))
        bars=ax.bar(['Con IRT2','Solo IRT1'],[s['N_irt2'],s['N_solo1']],color=[MID,'#CCC'],width=.5)
        for b,v in zip(bars,[s['N_irt2'],s['N_solo1']]):
            ax.text(b.get_x()+b.get_width()/2.,b.get_height()+.5,str(v),
                    ha='center',va='bottom',fontsize=11,fontweight='bold',color=NAVY)
        ax.set_title('Estado de seguimiento',fontsize=11,color=NAVY,fontweight='bold',pad=8)
        ax.set_facecolor('#F8FAFD');fig.patch.set_facecolor('#F8FAFD')
        ax.spines[['top','right','left']].set_visible(False);ax.yaxis.set_visible(False)
        plt.tight_layout();st.pyplot(fig);plt.close()

    with gc2:
        sv=[s['n_verde'],s['n_naranja'],s['n_rojo'],s['N_irt2']]
        sl=['<60d','60-89d','90+d','Completados']; sc=[GREEN,ORANGE,RED,MID]
        sv_f=[v for v in sv if v>0]; sl_f=[l for l,v in zip(sl,sv) if v>0]; sc_f=[c for c,v in zip(sc,sv) if v>0]
        fig,ax=plt.subplots(figsize=(4.5,3.2))
        if sv_f:
            w,_,at=ax.pie(sv_f,colors=sc_f,autopct='%1.0f%%',startangle=90,
                wedgeprops={'edgecolor':'white','linewidth':1.5},textprops={'fontsize':9})
            for a in at: a.set_color('white');a.set_fontweight('bold')
            ax.legend(w,[f'{l} ({v})' for l,v in zip(sl_f,sv_f)],
                loc='lower center',bbox_to_anchor=(.5,-.3),fontsize=7.5,ncol=2,frameon=False)
        ax.set_title('Semáforo seguimiento',fontsize=11,color=NAVY,fontweight='bold',pad=8)
        fig.patch.set_facecolor('#F8FAFD');plt.tight_layout();st.pyplot(fig);plt.close()

    with gc3:
        sust=s.get('sust_dist',{})
        sd=pd.DataFrame(list(sust.items()),columns=['S','n']).sort_values('n') if sust else pd.DataFrame()
        fig,ax=plt.subplots(figsize=(4.5,3.2))
        if not sd.empty:
            cols_=[MID if i%2==0 else ACCENT for i in range(len(sd))]
            ax.barh(sd['S'],sd['n'],color=cols_,height=.6)
            tot=sd['n'].sum()
            for b,v in zip(ax.patches,sd['n']):
                ax.text(b.get_width()+.3,b.get_y()+b.get_height()/2,
                        f'{v} ({round(v/tot*100,1) if tot else 0}%)',va='center',fontsize=8,color=NAVY)
            ax.spines[['top','right','bottom']].set_visible(False);ax.xaxis.set_visible(False)
        ax.set_title('Sustancia principal (IRT1)',fontsize=11,color=NAVY,fontweight='bold',pad=8)
        ax.set_facecolor('#F8FAFD');fig.patch.set_facecolor('#F8FAFD')
        plt.tight_layout();st.pyplot(fig);plt.close()

    with st.expander('📋 Log de procesamiento'):
        for log in R['logs']: st.text(log)

    # ── Descargas ─────────────────────────────────────────────────────────────
    st.markdown('---')
    st.markdown('<div class="sec">⬇️ Descargar reportes</div>', unsafe_allow_html=True)

    fname_base = os.path.splitext(st.session_state.get('filename','base'))[0]
    if fc: fname_base += f'_{fc}'
    today_str = datetime.now().strftime('%Y-%m-%d')
    outputs   = st.session_state.get('outputs',{})
    sel       = st.session_state.get('seleccion',{})

    d1,d2,d3=st.columns(3)
    with d1:
        st.markdown('<div class="outcard"><h4>📊 Base Wide completa</h4>'
                    '<p>Hojas: Wide · Resumen · Alertas · Calidad</p></div>', unsafe_allow_html=True)
        st.download_button('⬇️ Base Wide (.xlsx)',
            data=R['excel_bytes'].getvalue(),
            file_name=f'IRT_Base_Wide_{fname_base}_{today_str}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True, key='dl_wide')

    for key,col,dlkey in [('caract_excel',d2,'dl_ce'),('seg_excel',d3,'dl_se')]:
        o=outputs.get(key,{}); lbl,fmt,desc=LABELS[key]
        with col:
            st.markdown(f'<div class="outcard"><h4>{lbl}</h4><p>{desc}</p></div>',unsafe_allow_html=True)
            if not sel.get(key,False): st.caption('No seleccionado')
            elif o.get('ok'):
                st.download_button(f'⬇️ {fmt}',data=o['buf'].getvalue(),
                    file_name=o['fname'],mime=o['mime'],use_container_width=True,key=dlkey)
            else: st.warning(f"⚠️ {o.get('error','Error')[:100]}")

    st.markdown('---')
    d4,d5=st.columns(2)
    for key,col,dlkey in [('word_caract',d4,'dl_wc'),('word_seg',d5,'dl_ws')]:
        o=outputs.get(key,{}); lbl,fmt,desc=LABELS[key]
        with col:
            st.markdown(f'<div class="outcard"><h4>{lbl}</h4><p>{desc}</p></div>',unsafe_allow_html=True)
            if not sel.get(key,False): st.caption('No seleccionado')
            elif o.get('ok'):
                st.download_button(f'⬇️ {fmt}',data=o['buf'].getvalue(),
                    file_name=o['fname'],mime=o['mime'],use_container_width=True,key=dlkey)
            else: st.warning(f"⚠️ {o.get('error','Error')[:100]}")

    # ── Distribución por centros ───────────────────────────────────────────────
    if 'wide_path' in st.session_state and not filtro_centro_val:
        st.markdown('---')
        st.markdown('<div class="sec">📦 Distribución por centros</div>', unsafe_allow_html=True)
        st.markdown(
            '<div style="background:#EEF4FB;border-left:4px solid #2E75B6;'
            'padding:.8rem 1.2rem;border-radius:6px;margin-bottom:1rem;">'
            '<b>¿Qué genera este botón?</b><br>'
            'Un archivo <b>.zip</b> con una carpeta por cada centro. '
            'Cada carpeta incluye la base Wide filtrada + los reportes seleccionados.'
            '</div>', unsafe_allow_html=True)

        dc1,dc2 = st.columns(2)
        with dc1:
            d_ce = st.checkbox('📋 Excel caracterización', value=True, key='d_ce')
            d_se = st.checkbox('📋 Excel seguimiento',     value=True, key='d_se')
        with dc2:
            d_wc = st.checkbox('📄 Word caracterización',  value=True, key='d_wc')
            d_ws = st.checkbox('📄 Word seguimiento',      value=True, key='d_ws')

        keys_dist = [k for k,v in {'caract_excel':d_ce,'seg_excel':d_se,
                                    'word_caract':d_wc,'word_seg':d_ws}.items() if v]
        n_centros = len(centros_disponibles)
        st.caption(f'Se generarán **{n_centros} carpetas** — una por centro')

        if st.button('📦 Generar paquetes por centro', use_container_width=True, key='btn_dist'):
            wide_path_dist = st.session_state['wide_path']
            status_box = st.empty()
            prog_dist  = st.progress(0, text='Iniciando...')

            def _cb(i, total, centro):
                pct = i/total if total else 1
                txt = f'Centro {i+1}/{total}: {centro}' if centro!='listo' else '✅ ZIP generado'
                prog_dist.progress(pct, text=txt); status_box.info(txt)

            try:
                with st.spinner('Generando paquetes...'):
                    zip_buf = run_paquetes_centros(wide_path_dist, keys_sel=keys_dist, progress_cb=_cb)
                today_str2 = datetime.now().strftime('%Y-%m-%d')
                prog_dist.progress(1.0, text='✅ Listo')
                status_box.success(f'✅ ZIP con {n_centros} carpetas generado')
                st.download_button(
                    label=f'⬇️ Descargar ZIP ({n_centros} centros)',
                    data=zip_buf.getvalue(),
                    file_name=f'QALAT_IRT_Paquetes_{today_str2}.zip',
                    mime='application/zip',
                    use_container_width=True, key='dl_dist')
            except Exception as e:
                st.error(f'❌ Error: {e}')

if not uploaded and 'result' not in st.session_state:
    st.markdown("""<div style="text-align:center;padding:3rem;color:#888;">
        <div style="font-size:3rem;">📤</div>
        <div style="font-size:1.1rem;margin-top:1rem;">Sube tu Excel IRT para comenzar</div>
        <div style="font-size:.85rem;margin-top:.5rem;color:#aaa;">Base bruta exportada de Jotform · instrumento IRT</div>
    </div>""", unsafe_allow_html=True)
