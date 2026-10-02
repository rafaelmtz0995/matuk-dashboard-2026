#!/usr/bin/env python3
"""MATUK Dashboard 2026 - Regenera var DATA, BREAKDOWN, FDATA, COSTOS"""
import os, sys, datetime, json
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable,"-m","pip","install","openpyxl",
                           "--quiet","--break-system-packages"])
    import openpyxl

HOME = os.path.expanduser("~")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

FLUJO_FILE = None
for c in [
    os.path.join(HOME, "Matuk Automation service", "Hugo Carreon - BANCOS", "Flujo de Caja 2026.xlsx"),
    os.path.join(HOME, "mnt", "Matuk Automation service", "Hugo Carreon - BANCOS", "Flujo de Caja 2026.xlsx"),
]:
    if os.path.exists(c): FLUJO_FILE = c; break

PAGOS_FILE = None
for c in [
    os.path.join(HOME, "Matuk Automation service", "Matuk Automation Repository - Documentos", "Reporte de horas 2022", "Servicio Administrativo Pagos 2026.xlsx"),
    os.path.join(HOME, "mnt", "Matuk Automation service", "Matuk Automation Repository - Documentos", "Reporte de horas 2022", "Servicio Administrativo Pagos 2026.xlsx"),
    os.path.join(SCRIPT_DIR, "Servicio Administrativo Pagos 2026.xlsx"),
]:
    if os.path.exists(c): PAGOS_FILE = c; break

OUTPUT_HTML = os.path.join(SCRIPT_DIR, "dashboard.html")

print(f"FLUJO : {FLUJO_FILE}")
print(f"PAGOS : {PAGOS_FILE}")
print(f"HTML  : {OUTPUT_HTML}")

MESES = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP']
MES_MAP = {1:'ENE',2:'FEB',3:'MAR',4:'ABR',5:'MAY',6:'JUN',7:'JUL',8:'AGO',9:'SEP'}
CMONTHS = ['NOV','DIC','ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP']

MES_LARGO = {
    'ENERO':'ENE','FEBRERO':'FEB','MARZO':'MAR','MARZO ':'MAR','ENERO ':'ENE',
    'ABRIL':'ABR','MAYO':'MAY','JUNIO':'JUN','JULIO':'JUL','AGOSTO':'AGO','SEPTIEMBRE':'SEP',
    'FEBRERO ':'FEB','ABRIL ':'ABR','MAYO ':'MAY','JUNIO ':'JUN',
    'JULIO ':'JUL','AGOSTO ':'AGO','SEPTIEMBRE':'SEP','SEPTIEMBRE ':'SEP',
}

EXCLUIR_ING = {
    'COB','TRASPASO','COMPENSACION POR RETRASO','TOTAL GENERAL',
    'DEVOLUCION DEPOSITO EN GARANTIA','DEVOLUCION GUILLERMO FORTINO',
    'DEVOLUCION RAFAEL RODRIGUEZ','PAGO PRESTAMO GUILLERMO FORTINO',
    'PAGO PRESTAMO RAFAEL RODRIGUEZ','OTROS INGRESOS','MATUK LLC',
    'TOTAL COSTOS','COSTOS','TOTAL GASTOS','AMEX',
}

def sf(v):
    try: return float(v) if v is not None else 0.0
    except: return 0.0

# ── INGRESOS MXN ─────────────────────────────────────────────────────────────
def extraer_ingresos():
    print("  Ingresos MXN (Tabla )...")
    wb = openpyxl.load_workbook(FLUJO_FILE, read_only=True, data_only=True)
    ws = wb['Tabla ']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[1]
    col_map = {}
    for j,h in enumerate(header):
        if isinstance(h,str):
            k = h.strip().upper()
            if k in MES_LARGO and MES_LARGO[k] not in col_map:
                col_map[MES_LARGO[k]] = j
    clientes_mes = {m:{} for m in MESES}
    total_mes    = {m:0.0 for m in MESES}
    for row in rows[2:]:
        nombre = str(row[0]).strip().upper() if row[0] else ""
        if not nombre or nombre in EXCLUIR_ING: continue
        for m,col in col_map.items():
            if col < len(row):
                val = sf(row[col])
                if val > 0:
                    clientes_mes[m][nombre] = clientes_mes[m].get(nombre,0.0)+val
                    total_mes[m] += val
    wb.close()
    return clientes_mes, total_mes

# ── USD + TC ──────────────────────────────────────────────────────────────────
def extraer_usd_tc():
    print("  USD y TC (Res flujo enero-julio)...")
    wb = openpyxl.load_workbook(FLUJO_FILE, read_only=True, data_only=True)
    ws = wb['Res flujo enero-julio']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[3]
    col_map = {}
    for j,h in enumerate(header):
        if isinstance(h,str):
            k = h.strip().upper()
            if k in MES_LARGO and MES_LARGO[k] not in col_map:
                col_map[MES_LARGO[k]] = j
    ing_usd = {m:0.0 for m in MESES}
    tc_mes  = {m:0.0 for m in MESES}
    for row in rows:
        etiq = str(row[0]).strip().upper() if row[0] else ""
        if etiq == 'INGRESOS USD':
            for m,col in col_map.items():
                if col < len(row): ing_usd[m] = sf(row[col])
        elif etiq == 'TC CIERRE DE MES':
            for m,col in col_map.items():
                if col < len(row): tc_mes[m]  = sf(row[col])
    wb.close()
    return ing_usd, tc_mes

# ── FDATA desde FLUJO detalle ─────────────────────────────────────────────────
def extraer_fdata():
    print("  FDATA (FLUJO detalle)...")
    wb = openpyxl.load_workbook(FLUJO_FILE, read_only=True, data_only=True)
    ws = wb['FLUJO detalle']
    rows = list(ws.iter_rows(values_only=True))

    # Fila 5 (índice 5): nombres de meses. DLS=col(N-1), MX=col(N)
    mes_cols = {}
    header_row = rows[5]
    for j, cell in enumerate(header_row):
        if isinstance(cell, str):
            k = cell.strip().upper().rstrip()
            if k in MES_LARGO:
                m = MES_LARGO[k]
                if m in MESES and m not in mes_cols:
                    mes_cols[m] = (j-1, j)

    print(f"    mes_cols: {mes_cols}")

    def get(row, m, d):
        idx = mes_cols.get(m)
        if not idx: return 0.0
        col = idx[0] if d=='dls' else idx[1]
        return sf(row[col]) if col < len(row) else 0.0

    fdata = {m:{k:0.0 for k in [
        'ing_mx','ing_dls','egr_mx','egr_dls',
        'saldo_mx','saldo_dls','horas_mx','horas_dls',
        'viat_mx','viat_dls','nom_mx','nom_dls',
        'imp_mx','imp_dls','amex_mx','amex_dls','cf_mx','cf_dls'
    ]} for m in MESES}

    TARGETS = {
        'TOTAL INGRESOS': 'ing', 'TOTAL INGRESOS ': 'ing',
        'TOTAL EGRESOS':  'egr', 'TOTAL EGRESOS ': 'egr',
        'SALDO FINAL BANCOS': 'saldo', 'SALDO FINAL BANCOS ': 'saldo',
        'HORAS':   'horas', 'VIATICOS': 'viat', 'VIATICOS ': 'viat',
        'NOM ADMON': 'nom',  'NOM ADMON ': 'nom',
        'IMPUESTOS': 'imp',  'IMPUESTOS '  : 'imp',
        'AMEX': 'amex',
        'COSTOS FIJOS': 'cf', 'COSTOS FIJOS ': 'cf',
    }
    found = set()
    for row in rows:
        label = str(row[0]).strip().upper() if row[0] else ""
        if label in TARGETS and label not in found:
            key = TARGETS[label]
            found.add(label)
            for m in MESES:
                fdata[m][key+'_dls'] = get(row, m, 'dls')
                fdata[m][key+'_mx']  = get(row, m, 'mx')
    wb.close()

    print(f"    ENE ing_mx={fdata['ENE']['ing_mx']:,.0f}  ing_dls={fdata['ENE']['ing_dls']:,.0f}")
    return fdata

# ── HORAS detalle para DATA y BREAKDOWN ──────────────────────────────────────
def extraer_horas():
    print("  Horas detalle (hoja 2026)...")
    wb = openpyxl.load_workbook(PAGOS_FILE, read_only=True, data_only=True)
    ws = wb['2026']
    data   = {m:{'hours':0.0,'st':0.0,'ot':0.0,'dt':0.0,'res':set()} for m in MESES}
    bd_tmp = {m:{} for m in MESES}

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[3] or 'HOURS' not in str(row[3]).upper(): continue
        wm = row[11]
        if not isinstance(wm, datetime.datetime) or wm.year != 2026: continue
        if not (1 <= wm.month <= 9): continue
        m = MES_MAP[wm.month]

        hrs=sf(row[13]); st=sf(row[14]); ot=sf(row[15]); dt=sf(row[16])
        name = str(row[2]).strip() if row[2] else ""
        svc  = str(row[4]).strip() if row[4] else ""
        po   = str(row[5]).strip() if row[5] else ""
        cust = str(row[6]).strip().upper() if row[6] else ""

        data[m]['hours']+=hrs; data[m]['st']+=st; data[m]['ot']+=ot; data[m]['dt']+=dt
        if name: data[m]['res'].add(name)

        key = (svc,po,cust)
        if key not in bd_tmp[m]:
            bd_tmp[m][key]={'s':svc,'po':po,'c':cust,'st':0,'ot':0,'dt':0,'r':set()}
        bd_tmp[m][key]['st']+=st; bd_tmp[m][key]['ot']+=ot; bd_tmp[m][key]['dt']+=dt
        if name: bd_tmp[m][key]['r'].add(name)

    breakdown={}
    for m in MESES:
        items=[{'s':v['s'],'po':v['po'],'c':v['c'],
                'st':round(v['st'],2),'ot':round(v['ot'],2),'dt':round(v['dt'],2),
                't':round(v['st']+v['ot']+v['dt'],2),'r':len(v['r'])}
               for v in bd_tmp[m].values()]
        breakdown[m]=sorted(items,key=lambda x:-x['t'])
    wb.close()
    return data, breakdown

# ── COSTOS desde hoja COSTOS + by_cust desde hoja 2026 ───────────────────────
def extraer_costos(tc_mes):
    print("  Costos mensuales (hoja COSTOS + hoja 2026)...")
    wb = openpyxl.load_workbook(PAGOS_FILE, read_only=True, data_only=True)

    # 1. Resumen mensual de hoja COSTOS (filas 10-22)
    ws = wb['COSTOS']
    rows = list(ws.iter_rows(values_only=True))
    costos = {}
    MES_KEY = {  # fecha.month para 2025/2026
        (2025,11):'NOV', (2025,12):'DIC',
        (2026,1):'ENE',  (2026,2):'FEB',  (2026,3):'MAR',
        (2026,4):'ABR',  (2026,5):'MAY',  (2026,6):'JUN',
        (2026,7):'JUL',  (2026,8):'AGO',  (2026,9):'SEP',
    }
    for row in rows[9:24]:  # filas con datos mensuales
        fecha = row[1]
        if not isinstance(fecha, datetime.datetime): continue
        key = (fecha.year, fecha.month)
        if key not in MES_KEY: continue
        m = MES_KEY[key]
        costos[m] = {
            'usd_hrs':   sf(row[7]),   # Costo Hrs USD
            'usd_total': sf(row[5]),   # Total USD
            'mxn_hrs':   sf(row[7]),   # Costo Hrs MXN  (usamos col7=costo_hrs para horas)
            'mxn_total': sf(row[8]),   # Total MXN
            'by_cust':   [],
        }
        # Corregir: col indices del resumen mensual son:
        # col1=fecha, col2=horas_usd, col3=costo_hrs_usd, col4=perdiem_usd, col5=total_usd
        # col6=horas_mxn, col7=costo_mxn, col8=total_mxn
        costos[m] = {
            'usd_hrs':   sf(row[3]),   # Costo Hrs USD
            'usd_total': sf(row[5]),   # Total USD
            'mxn_hrs':   sf(row[7]),   # Costo Hrs MXN
            'mxn_total': sf(row[8]),   # Total MXN
            'by_cust':   [],
        }

    # 2. by_cust desde hoja 2026 — col28=Costo Hrs TOTAL USD, col30=Costo TOTAL USD
    ws2 = wb['2026']
    by_cust_tmp = {m:{} for m in CMONTHS}

    for row in ws2.iter_rows(min_row=2, values_only=True):
        if not row[3] or 'HOURS' not in str(row[3]).upper(): continue
        wm = row[11]
        if not isinstance(wm, datetime.datetime): continue
        key = (wm.year, wm.month)
        if key not in MES_KEY: continue
        m = MES_KEY[key]

        cust = str(row[6]).strip().upper() if row[6] else "SIN CLIENTE"
        po   = str(row[5]).strip() if row[5] else ""
        svc  = str(row[4]).strip() if row[4] else ""
        divisa = str(row[19]).strip().upper() if row[19] else "USA"
        hrs_total = sf(row[13])  # WORK HOURS
        costo_hrs_usd  = sf(row[28]) if len(row)>28 else 0.0  # Costo Hrs TOTAL USD
        costo_total_usd= sf(row[30]) if len(row)>30 else 0.0  # Costo TOTAL USD

        if cust not in by_cust_tmp[m]:
            by_cust_tmp[m][cust] = {'usd_hrs':0,'usd':0,'mxn_hrs':0,'mxn':0,'by_po':{}}
        c = by_cust_tmp[m][cust]
        if divisa == 'MXN':
            tc = tc_mes.get(m, 18.5) if m in MESES else 18.5
            c['mxn_hrs'] += costo_hrs_usd
            c['mxn']     += costo_total_usd
        else:
            c['usd_hrs'] += costo_hrs_usd
            c['usd']     += costo_total_usd

        po_key = po
        if po_key not in c['by_po']:
            c['by_po'][po_key] = {'po':po,'usd_hrs':0,'usd':0,'mxn_hrs':0,'mxn':0}
        p = c['by_po'][po_key]
        if divisa == 'MXN':
            p['mxn_hrs'] += costo_hrs_usd; p['mxn'] += costo_total_usd
        else:
            p['usd_hrs'] += costo_hrs_usd; p['usd'] += costo_total_usd

    # Armar estructura final by_cust
    for m in CMONTHS:
        if m not in costos:
            costos[m] = {'usd_hrs':0,'usd_total':0,'mxn_hrs':0,'mxn_total':0,'by_cust':[]}
        bycust = []
        for cname, cv in sorted(by_cust_tmp[m].items(), key=lambda x:-x[1]['usd']):
            by_po = [{'po':pv['po'],'usd_hrs':round(pv['usd_hrs'],2),'usd':round(pv['usd'],2),
                      'mxn_hrs':round(pv['mxn_hrs'],2),'mxn':round(pv['mxn'],2)}
                     for pv in sorted(cv['by_po'].values(),key=lambda x:-x['usd'])]
            bycust.append({'c':cname,
                           'usd_hrs':round(cv['usd_hrs'],2),'usd':round(cv['usd'],2),
                           'mxn_hrs':round(cv['mxn_hrs'],2),'mxn':round(cv['mxn'],2),
                           'by_po':by_po})
        costos[m]['by_cust'] = bycust
        costos[m]['tc'] = round(tc_mes.get(m, 0) if m in MESES else 0, 4)

    wb.close()

    # Ajustar usd_hrs = suma de by_cust cuando hoja COSTOS tiene 0
    for m in CMONTHS:
        if costos[m]['usd_hrs'] == 0 and costos[m]['by_cust']:
            costos[m]['usd_hrs'] = round(sum(c['usd_hrs'] for c in costos[m]['by_cust']),2)
            costos[m]['usd_total'] = round(sum(c['usd'] for c in costos[m]['by_cust']),2)

    print(f"    Meses con datos: {[m for m in CMONTHS if costos[m]['usd_total']>0]}")
    return costos

# ── REEMPLAZAR VARIABLE EN HTML ───────────────────────────────────────────────
def _replace_var(html, varname, new_value_str):
    pos = html.find(f'var {varname}')
    if pos < 0:
        print(f"  WARN: no se encontró 'var {varname}'")
        return html
    i = html.find('{', pos)
    if i < 0: return html
    depth = 0
    while i < len(html):
        if html[i] == '{': depth += 1
        elif html[i] == '}':
            depth -= 1
            if depth == 0:
                end = i+1
                j = end
                while j < len(html) and html[j] in ' \t': j+=1
                if j < len(html) and html[j] == ';': end = j+1
                break
        i += 1
    else:
        return html
    html = html[:pos] + f"var {varname}={new_value_str};" + html[end:]
    print(f"  OK var {varname}")
    return html

MARCA_INI = "// ─── AUTO-GENERADO POR regenerar_dashboard_v2.py"
MARCA_FIN = "// ─── FIN DATOS AUTO-GENERADOS"

def actualizar_html(clientes_mes, total_mes, ing_usd, tc_mes,
                    horas_data, breakdown, costos, fdata):
    if not os.path.exists(OUTPUT_HTML):
        print(f"ERROR: {OUTPUT_HTML}"); sys.exit(1)

    print("  Leyendo dashboard.html...")
    with open(OUTPUT_HTML,'r',encoding='utf-8') as f: html=f.read()
    print(f"  Tamaño: {len(html):,} bytes")

    # var DATA
    data_js = {m:{'hours':round(horas_data[m]['hours'],0),'st':round(horas_data[m]['st'],0),
                  'ot':round(horas_data[m]['ot'],0),'dt':round(horas_data[m]['dt'],0),
                  'res':len(horas_data[m]['res'])} for m in MESES}
    html = _replace_var(html, 'DATA', json.dumps(data_js))

    # var BREAKDOWN
    html = _replace_var(html, 'BREAKDOWN', json.dumps(breakdown, ensure_ascii=False))

    # var FDATA
    fdata_r = {m:{k:round(v,2) for k,v in row.items()} for m,row in fdata.items()}
    html = _replace_var(html, 'FDATA', json.dumps(fdata_r, ensure_ascii=False))

    # var COSTOS
    costos_r = {}
    for m in CMONTHS:
        if m in costos:
            d = costos[m]
            costos_r[m] = {
                'usd_hrs':   round(d['usd_hrs'],2),
                'usd_total': round(d['usd_total'],2),
                'mxn_hrs':   round(d['mxn_hrs'],2),
                'mxn_total': round(d['mxn_total'],2),
                'by_cust':   d['by_cust'],
                'tc':        d.get('tc',0),
            }
    html = _replace_var(html, 'COSTOS', json.dumps(costos_r, ensure_ascii=False))

    # Bloque marcadores
    idx_i = html.find(MARCA_INI)
    idx_f = html.find(MARCA_FIN)
    if idx_i >= 0 and idx_f >= 0:
        idx_f_end = idx_f + len(MARCA_FIN)
        while idx_f_end < len(html) and html[idx_f_end] not in '\n': idx_f_end+=1
        ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ytd = defaultdict(float)
        for m in MESES:
            for cli,val in clientes_mes[m].items(): ytd[cli]+=val
        ytd_sorted = dict(sorted(ytd.items(),key=lambda x:-x[1]))

        L = [f"// Generado: {ts}", ""]
        L.append("var FCLIENTES = "+json.dumps({k:{"mx":round(v,0)} for k,v in ytd_sorted.items()},ensure_ascii=False)+";")
        L.append("var FCLIENTES_MES = "+json.dumps({m:{k:round(v,0) for k,v in d.items()} for m,d in clientes_mes.items()},ensure_ascii=False)+";")
        L.append("var FING_MES = "+json.dumps({m:round(total_mes[m],2) for m in MESES})+";")
        L.append("var FING_USD_MES = "+json.dumps({m:round(ing_usd[m],2) for m in MESES})+";")
        L.append("var FTC_MES = "+json.dumps({m:round(tc_mes[m],4) for m in MESES})+";")
        L.append("var FHORAS_MES = "+json.dumps({m:round(horas_data[m]['hours'],2) for m in MESES})+";")
        L.append("var FHORAS_CLI_MES = "+json.dumps({m:{} for m in MESES})+";")
        L.append("var FHORAS_SVC_MES = "+json.dumps({m:{} for m in MESES})+";")
        L.append("var FRECURSOS_MES = "+json.dumps({m:len(horas_data[m]['res']) for m in MESES})+";")
        costos_mes_simple = {}
        for m in MESES:
            if m in costos:
                d=costos[m]
                costos_mes_simple[m]={'horas_usd':round(d['usd_hrs'],2),
                                      'total_usd':round(d['usd_total'],2),
                                      'total_mxn':round(d['mxn_total'],2)}
            else:
                costos_mes_simple[m]={'horas_usd':0,'total_usd':0,'total_mxn':0}
        L.append("var FCOSTOS_MES = "+json.dumps(costos_mes_simple,ensure_ascii=False)+";")

        bloque = MARCA_INI+"\n"+"\n".join(L)+"\n"+MARCA_FIN
        html  = html[:idx_i]+bloque+html[idx_f_end:]
        print("  OK bloque marcadores")
    else:
        print("  WARN: marcadores no encontrados")

    with open(OUTPUT_HTML,'w',encoding='utf-8') as f: f.write(html)
    print(f"  OK guardado ({len(html):,} bytes)")



# ── FIXES POST-REGENERACION ───────────────────────────────────────────────────
def aplicar_fixes(html):
    """Re-aplica todos los fixes manuales que el regenerador sobreescribe."""
    fixes_ok = []
    fixes_fail = []

    # FIX 1: showTab — usar 'block' en vez de '' para tabs ocultas por CSS
    OLD1 = """function showTab(t){
  document.getElementById('view-resumen').style.display = t==='resumen' ? '' : 'none';
  document.getElementById('view-horas').style.display = t==='horas' ? '' : 'none';
  document.getElementById('view-costos').style.display = t==='costos' ? '' : 'none';
  document.getElementById('view-flujo').style.display = t==='flujo' ? '' : 'none';"""
    NEW1 = """function showTab(t){
  document.getElementById('view-resumen').style.display = t==='resumen' ? 'block' : 'none';
  document.getElementById('view-horas').style.display = t==='horas' ? 'block' : 'none';
  document.getElementById('view-costos').style.display = t==='costos' ? 'block' : 'none';
  document.getElementById('view-flujo').style.display = t==='flujo' ? 'block' : 'none';"""
    if OLD1 in html:
        html = html.replace(OLD1, NEW1); fixes_ok.append("showTab block")
    elif NEW1 in html:
        fixes_ok.append("showTab block (ya OK)")
    else:
        fixes_fail.append("showTab block")

    # FIX 2: KPI Costo Operativo en Resumen (5to KPI)
    OLD2 = "    {label:'Horas Facturadas YTD',val:horasTotal.toLocaleString('es-MX')+' hrs', sub:'ST: '+stTotal.toLocaleString('es-MX')+' · OT: '+otTotal.toLocaleString('es-MX'),accent:'#7C3AED', trend:''},"
    NEW2 = """    (function(){ const costoOpTotal = months.reduce(function(a,m){ return a + (FDATA[m]?(FDATA[m].horas_mx||0)+(FDATA[m].viat_mx||0)+(FDATA[m].nom_mx||0)+(FDATA[m].imp_mx||0):0); }, 0); const pct = egrTotal>0?(costoOpTotal/egrTotal*100).toFixed(1)+'% de egresos totales':''; return {label:'Costo Operativo', val:fmt(costoOpTotal), sub:'Horas + Viát + Nóm + Imp<br><span style="color:#7C3AED;font-weight:700">'+pct+'</span>', accent:'#7C3AED', trend:''}; })(),"""
    if OLD2 in html:
        html = html.replace(OLD2, NEW2); fixes_ok.append("Costo Operativo KPI")
    elif 'Costo Operativo' in html:
        fixes_ok.append("Costo Operativo KPI (ya OK)")
    else:
        fixes_fail.append("Costo Operativo KPI")

    # FIX 3: cFilterMes llama buildCostosCharts
    OLD3 = """function cFilterMes(m){
  cCurrentMes = m;
  document.querySelectorAll('.cfilter-btn').forEach(b => b.classList.toggle('active', b.dataset.m === m));
  // Clear search if active
  const inp = document.getElementById('cSearchInput');
  if(inp && inp.value){ inp.value=''; searchCostosTable(''); }
  buildCostosTable();
}"""
    NEW3 = """function cFilterMes(m){
  cCurrentMes = m;
  document.querySelectorAll('.cfilter-btn').forEach(b => b.classList.toggle('active', b.dataset.m === m));
  // Clear search if active
  const inp = document.getElementById('cSearchInput');
  if(inp && inp.value){ inp.value=''; searchCostosTable(''); }
  buildCostosTable();
  buildCostosCharts();
}

// ---- GRAFICAS COSTOS ----
var _costosTendChart = null;
var _costosBarChart = null;

function buildCostosCharts(){
  var allMonths = ['NOV','DIC','ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT'];
  var meses = allMonths.filter(function(m){ return COSTOS[m] && ((COSTOS[m].usd_total||0)+(COSTOS[m].mxn_total||0)) > 0; });
  var selectedMes = (typeof cCurrentMes!=='undefined') ? cCurrentMes : 'TODOS';
  var monthlyCosts = meses.map(function(m){ return (COSTOS[m].usd_total||0) + (COSTOS[m].mxn_total||0)/17; });
  var acum = 0;
  var acumCosts = monthlyCosts.map(function(v){ acum += v; return Math.round(acum); });
  var monthlyRound = monthlyCosts.map(function(v){ return Math.round(v); });
  var labels = meses.map(function(m){ return (m==='NOV'||m==='DIC') ? m+' 25' : m; });
  var ctx1 = document.getElementById('costos-tendencia-chart');
  if(!ctx1) return;
  if(_costosTendChart){ _costosTendChart.destroy(); _costosTendChart=null; }
  _costosTendChart = new Chart(ctx1, {
    data: { labels: labels, datasets: [
      { type:'line', label:'Acumulado', data:acumCosts, borderColor:'#1F5BA6', backgroundColor:'rgba(31,91,166,.08)', borderWidth:2, pointRadius:4, pointBackgroundColor:'#fff', pointBorderColor:'#1F5BA6', fill:true, yAxisID:'y1', tension:0.3 },
      { type:'line', label:'Mensual', data:monthlyRound, borderColor:'#F59E0B', borderWidth:2, borderDash:[6,4], pointRadius:4, pointBackgroundColor:'#fff', pointBorderColor:'#F59E0B', fill:false, yAxisID:'y2', tension:0.3 }
    ]},
    options: { responsive:true, interaction:{mode:'index',intersect:false},
      plugins:{ legend:{position:'bottom',labels:{font:{size:11},boxWidth:14}}, tooltip:{callbacks:{label:function(c){ return c.dataset.label+': $'+c.parsed.y.toLocaleString('en-US'); }}} },
      scales:{ x:{grid:{color:'#EEF2F7'},ticks:{color:'#6B7A99',font:{size:10}}},
        y1:{position:'left',grid:{color:'#EEF2F7'},ticks:{color:'#6B7A99',font:{size:10},callback:function(v){ return '$'+(v>=1000?(v/1000).toFixed(0)+'K':v); }}},
        y2:{position:'right',grid:{drawOnChartArea:false},ticks:{color:'#F59E0B',font:{size:10},callback:function(v){ return '$'+(v>=1000?(v/1000).toFixed(0)+'K':v); }}}
      }
    }
  });
  var ctx2 = document.getElementById('costos-barras-chart');
  if(!ctx2) return;
  if(_costosBarChart){ _costosBarChart.destroy(); _costosBarChart=null; }
  var barColors = meses.map(function(m){
    if(selectedMes === 'TODOS') return '#1F5BA6';
    return m === selectedMes ? '#1F5BA6' : 'rgba(31,91,166,.18)';
  });
  _costosBarChart = new Chart(ctx2, {
    type:'bar',
    data:{ labels:labels, datasets:[{ label:'Costo Total USD', data:monthlyRound, backgroundColor:barColors, borderRadius:4, borderSkipped:false }] },
    options:{ responsive:true, plugins:{legend:{display:false}, tooltip:{callbacks:{label:function(c){ return '$'+c.parsed.y.toLocaleString('en-US'); }}}},
      scales:{ x:{grid:{color:'#EEF2F7'},ticks:{color:'#6B7A99',font:{size:10}}}, y:{grid:{color:'#EEF2F7'},ticks:{color:'#6B7A99',font:{size:10},callback:function(v){ return '$'+(v>=1000?(v/1000).toFixed(0)+'K':v); }}} }
    }
  });
}"""
    if OLD3 in html:
        html = html.replace(OLD3, NEW3); fixes_ok.append("cFilterMes + buildCostosCharts JS")
    elif 'buildCostosCharts' in html:
        fixes_ok.append("buildCostosCharts (ya OK)")
    else:
        fixes_fail.append("cFilterMes + buildCostosCharts JS")

    # FIX 4: showTab costos llama buildCostosCharts
    OLD4 = "if(t==='costos') showPasivoMes('jul-26', document.getElementById('pbtn-jul'));"
    NEW4 = "if(t==='costos'){ showPasivoMes('jul-26', document.getElementById('pbtn-jul')); buildCostosCharts(); }"
    if OLD4 in html:
        html = html.replace(OLD4, NEW4); fixes_ok.append("showTab costos -> buildCostosCharts")
    elif NEW4 in html:
        fixes_ok.append("showTab costos (ya OK)")
    else:
        fixes_fail.append("showTab costos -> buildCostosCharts")

    # FIX 5: Canvas HTML para graficas costos (antes de HISTORIAL PASIVOS)
    CANVAS = """
      <!-- GRAFICAS COSTOS -->
      <div style="display:flex;gap:20px;padding:0 32px 32px 32px;box-sizing:border-box">
        <div style="flex:1;background:#fff;border-radius:12px;padding:20px;box-shadow:0 1px 6px rgba(0,0,0,.07)">
          <div style="font-size:.72rem;font-weight:800;color:#1F3A5F;text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px">
            <span style="color:#F59E0B;margin-right:6px">&#9679;</span>Tendencia Mensual vs Acumulado
          </div>
          <canvas id="costos-tendencia-chart" height="110"></canvas>
        </div>
        <div style="flex:1;background:#fff;border-radius:12px;padding:20px;box-shadow:0 1px 6px rgba(0,0,0,.07)">
          <div style="font-size:.72rem;font-weight:800;color:#1F3A5F;text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px">
            <span style="color:#1F5BA6;margin-right:6px">&#9679;</span>Total de Costos por Mes (USD)
          </div>
          <canvas id="costos-barras-chart" height="110"></canvas>
        </div>
      </div>
"""
    if 'costos-tendencia-chart' not in html:
        anchor = '<!-- HISTORIAL PASIVOS SECTION -->'
        if anchor in html:
            html = html.replace(anchor, CANVAS + '\n' + anchor)
            fixes_ok.append("Canvas graficas costos insertado")
        else:
            fixes_fail.append("Canvas graficas costos (anchor no encontrado)")
    else:
        fixes_ok.append("Canvas graficas costos (ya OK)")

    print("  FIXES OK:", ", ".join(fixes_ok))
    if fixes_fail:
        print("  FIXES FALLIDOS:", ", ".join(fixes_fail))
    return html

def main():
    print("="*60)
    print("  MATUK Dashboard 2026 - Regenerando")
    print("="*60); print()

    if not FLUJO_FILE: print("ERROR: Flujo de Caja 2026.xlsx no encontrado"); sys.exit(1)
    if not PAGOS_FILE: print("ERROR: Pagos 2026.xlsx no encontrado"); sys.exit(1)
    if not os.path.exists(OUTPUT_HTML): print(f"ERROR: {OUTPUT_HTML}"); sys.exit(1)

    clientes_mes, total_mes = extraer_ingresos()
    ing_usd, tc_mes         = extraer_usd_tc()
    fdata                   = extraer_fdata()
    horas_data, breakdown   = extraer_horas()
    costos                  = extraer_costos(tc_mes)

    print()
    for m in MESES:
        d=horas_data[m]
        c=costos.get(m,{})
        print(f"  {m}: {round(d['hours'])}h  ING_MXN={total_mes[m]:>12,.0f}  COSTOS_USD={c.get('usd_total',0):>9,.0f}")
    print()

    actualizar_html(clientes_mes, total_mes, ing_usd, tc_mes,
                    horas_data, breakdown, costos, fdata)

    # ── Agregar meses futuros con ceros para evitar errores JS ────────────────
    # MONTHS, CMONTHS y FMONTHS en el HTML pueden tener SEP/OCT
    # DATA, FDATA, COSTOS, BREAKDOWN deben tener todos los meses de sus arrays
    print("  Agregando meses futuros con ceros...")
    with open(OUTPUT_HTML, 'r', encoding='utf-8') as f: html2 = f.read()
    changed = False

    def add_missing(html, varname, missing_months, empty_val):
        p = html.find('var ' + varname + '=')
        if p < 0: return html, False
        i = html.find('{', p)
        depth = 0
        j = i
        while j < len(html):
            if html[j] == '{': depth += 1
            elif html[j] == '}':
                depth -= 1
                if depth == 0:
                    obj = json.loads(html[i:j+1])
                    added = False
                    for m in missing_months:
                        if m not in obj:
                            obj[m] = empty_val(m)
                            added = True
                    if added:
                        new_str = 'var ' + varname + '=' + json.dumps(obj, ensure_ascii=False) + ';'
                        return html[:p] + new_str + html[j+1:], True
                    return html, False
            j += 1
        return html, False

    MESES_EXTRA = ['OCT','NOV','DIC']
    EMPTY_DATA   = lambda m: {'hours':0,'st':0,'ot':0,'dt':0,'res':0}
    EMPTY_FDATA  = lambda m: {k:0.0 for k in ['ing_mx','ing_dls','egr_mx','egr_dls',
                              'saldo_mx','saldo_dls','horas_mx','horas_dls','viat_mx','viat_dls',
                              'nom_mx','nom_dls','imp_mx','imp_dls','amex_mx','amex_dls','cf_mx','cf_dls']}
    EMPTY_BD     = lambda m: []

    # Get COSTOS sample structure for empty
    p_c = html2.find('var COSTOS=')
    i_c = html2.find('{', p_c)
    dep = 0; jc = i_c
    while jc < len(html2):
        if html2[jc] == '{': dep += 1
        elif html2[jc] == '}':
            dep -= 1
            if dep == 0: break
        jc += 1
    costos_obj = json.loads(html2[i_c:jc+1])
    sample_c = costos_obj.get('AGO', costos_obj.get('JUL', {}))
    EMPTY_COSTOS = lambda m: {k: (0 if isinstance(v,(int,float)) else ({} if isinstance(v,dict) else [])) for k,v in sample_c.items()}

    html2, c1 = add_missing(html2, 'DATA', MESES_EXTRA, EMPTY_DATA)
    html2, c2 = add_missing(html2, 'FDATA', MESES_EXTRA, EMPTY_FDATA)
    html2, c3 = add_missing(html2, 'BREAKDOWN', MESES_EXTRA, EMPTY_BD)
    html2, c4 = add_missing(html2, 'COSTOS', MESES_EXTRA, EMPTY_COSTOS)

    if c1 or c2 or c3 or c4:
        with open(OUTPUT_HTML, 'w', encoding='utf-8') as f: f.write(html2)
        print(f"  OK meses extra agregados — {len(html2):,} bytes")
    else:
        print("  OK todos los meses ya presentes")

    # Aplicar todos los fixes manuales post-regeneracion
    print("  Aplicando fixes post-regeneracion...")
    with open(OUTPUT_HTML, 'r', encoding='utf-8') as f: html_fix = f.read()
    html_fix = aplicar_fixes(html_fix)
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f: f.write(html_fix)

    print()
    print("="*60)
    print("  LISTO — Recarga el dashboard.html en el navegador")
    print("="*60)

if __name__ == "__main__":
    main()
