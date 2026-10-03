#!/usr/bin/env python3
"""
regenerar_dashboard_v2.py  —  MATUK Automation Dashboard 2026
Lee SOLO:
  1. Flujo de Caja 2026.xlsx   -> FLUJO detalle (FDATA, saldos, egresos)
                                  Tabla  (ingresos MXN por cliente)
                                  Res flujo enero-julio (USD + TC)
  2. Servicio Administrativo Pagos 2026.xlsx
                                -> hoja 2026 (DATA, BREAKDOWN de horas)
                                   hoja COSTOS (FCOSTOS_MES)

Reemplaza en dashboard.html:
  - var DATA = {...};          (totales de horas por mes)
  - var BREAKDOWN = {...};     (detalle horas por servicio/PO/cliente)
  - var COSTOS = {...};        (detalle costos USD/MXN por mes para tab Costos)
  - var FDATA = {...};         (flujo de caja completo para tab Flujo)
  - bloque entre marcadores   (FCLIENTES, FING_MES, FCOSTOS_MES, etc.)
"""
import os, sys, datetime, json, re
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable,"-m","pip","install","openpyxl","--quiet"])
    import openpyxl

# ── RUTAS ────────────────────────────────────────────────────────────────────
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
RAIZ        = os.path.join(SCRIPT_DIR, "..", "..")

FLUJO_FILE  = os.path.join(RAIZ, "Hugo Carreon - BANCOS", "Flujo de Caja 2026.xlsx")
PAGOS_FILE  = os.path.join(SCRIPT_DIR, "Servicio Administrativo Pagos 2026.xlsx")
OUTPUT_HTML = os.path.join(os.path.expanduser("~"), "mnt", "Desktop",
                           "MATUK Dashboard 2026", "dashboard.html")

# ── CONSTANTES ───────────────────────────────────────────────────────────────
MESES = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO']
MES_IDX = {m:i for i,m in enumerate(MESES)}

MES_LARGO = {
    'ENERO':'ENE','FEBRERO':'FEB','MARZO':'MAR','MARZO ':'MAR','ENERO ':'ENE',
    'ABRIL':'ABR','MAYO':'MAY','JUNIO':'JUN','JULIO':'JUL','AGOSTO':'AGO',
    'FEBRERO ':'FEB','ABRIL ':'ABR','MAYO ':'MAY','JUNIO ':'JUN',
    'JULIO ':'JUL','AGOSTO ':'AGO',
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

# ── 1. INGRESOS MXN por cliente (Tabla ) ─────────────────────────────────────
def extraer_ingresos():
    print("  Ingresos MXN por cliente (Tabla )...")
    wb = openpyxl.load_workbook(os.path.abspath(FLUJO_FILE), read_only=True, data_only=True)
    ws = wb['Tabla ']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[1]
    col_map = {MES_LARGO[h.strip().upper()]: j
               for j, h in enumerate(header)
               if isinstance(h,str) and h.strip().upper() in MES_LARGO}
    clientes_mes = {m:{} for m in MESES}
    total_mes    = {m:0.0 for m in MESES}
    for row in rows[2:]:
        nombre = str(row[0]).strip().upper() if row[0] else ""
        if not nombre or nombre in EXCLUIR_ING: continue
        for mes,col in col_map.items():
            val = sf(row[col])
            if val > 0:
                clientes_mes[mes][nombre] = clientes_mes[mes].get(nombre,0.0)+val
                total_mes[mes] += val
    wb.close()
    return clientes_mes, total_mes

# ── 2. USD y TC (Res flujo enero-julio) ──────────────────────────────────────
def extraer_usd_tc():
    print("  USD e TC (Res flujo enero-julio)...")
    wb = openpyxl.load_workbook(os.path.abspath(FLUJO_FILE), read_only=True, data_only=True)
    ws = wb['Res flujo enero-julio']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[3]
    col_map = {MES_LARGO[h.strip().upper()]: j
               for j, h in enumerate(header)
               if isinstance(h,str) and h.strip().upper() in MES_LARGO}
    ing_usd = {m:0.0 for m in MESES}
    tc_mes  = {m:0.0 for m in MESES}
    for row in rows:
        etiq = str(row[0]).strip().upper() if row[0] else ""
        if etiq == 'INGRESOS USD':
            for m,col in col_map.items(): ing_usd[m] = sf(row[col])
        elif etiq == 'TC CIERRE DE MES':
            for m,col in col_map.items(): tc_mes[m]  = sf(row[col])
    wb.close()
    return ing_usd, tc_mes

# ── 3. FDATA y COSTOS (FLUJO detalle) ────────────────────────────────────────
def extraer_flujo_detalle():
    print("  FDATA y COSTOS (FLUJO detalle)...")
    wb = openpyxl.load_workbook(os.path.abspath(FLUJO_FILE), read_only=True, data_only=True)
    ws = wb['FLUJO detalle']
    rows = list(ws.iter_rows(values_only=True))

    # Fila 5 (idx) tiene nombres de meses; columnas DLS/MX alternan cada 3 (None,dls,mx)
    header_row = rows[5]
    # Mapear mes -> (col_dls, col_mx)
    # El patrón es: col 2,3=ENE(dls,mx), 5,6=FEB, 8,9=MAR, ... salto de 3
    mes_cols = {}
    col = 2
    for cell in header_row[2:]:
        if isinstance(cell, str):
            key = cell.strip().upper().rstrip()
            if key in MES_LARGO:
                m = MES_LARGO[key]
                if m in MESES and m not in mes_cols:
                    mes_cols[m] = (col, col+1)
        col += 1

    def get(row, m, divisa):
        idx = mes_cols.get(m)
        if not idx: return 0.0
        return sf(row[idx[0] if divisa=='dls' else idx[1]])

    # Inicializar FDATA
    fdata = {m:{
        'ing_mx':0,'ing_dls':0,'egr_mx':0,'egr_dls':0,
        'saldo_mx':0,'saldo_dls':0,
        'horas_mx':0,'horas_dls':0,
        'viat_mx':0,'viat_dls':0,
        'nom_mx':0,'nom_dls':0,
        'imp_mx':0,'imp_dls':0,
        'amex_mx':0,'amex_dls':0,
        'cf_mx':0,'cf_dls':0,
    } for m in MESES}

    # COSTOS para la tab de Costos (estructura por mes con by_cust)
    costos_horas = {m:{'usd_hrs':0,'usd_total':0,'mxn_hrs':0,'mxn_total':0,'by_cust':[]} for m in MESES}

    for row in rows:
        label = str(row[0]).strip().upper() if row[0] else ""
        if not label: continue

        if label == 'TOTAL INGRESOS':
            for m in MESES:
                fdata[m]['ing_mx']  = get(row,m,'mx')
                fdata[m]['ing_dls'] = get(row,m,'dls')
        elif label == 'TOTAL EGRESOS ':
            for m in MESES:
                fdata[m]['egr_mx']  = get(row,m,'mx')
                fdata[m]['egr_dls'] = get(row,m,'dls')
        elif label == 'SALDO FINAL BANCOS ':
            for m in MESES:
                fdata[m]['saldo_mx']  = get(row,m,'mx')
                fdata[m]['saldo_dls'] = get(row,m,'dls')
        elif label == 'HORAS':
            for m in MESES:
                fdata[m]['horas_mx']  = get(row,m,'mx')
                fdata[m]['horas_dls'] = get(row,m,'dls')
                costos_horas[m]['usd_hrs']   = get(row,m,'dls')
                costos_horas[m]['usd_total'] = get(row,m,'dls')
        elif label == 'VIATICOS':
            for m in MESES:
                fdata[m]['viat_mx']  = get(row,m,'mx')
                fdata[m]['viat_dls'] = get(row,m,'dls')
        elif label == 'NOM ADMON ':
            for m in MESES:
                fdata[m]['nom_mx']  = get(row,m,'mx')
                fdata[m]['nom_dls'] = get(row,m,'dls')
        elif label == 'IMPUESTOS ':
            for m in MESES:
                fdata[m]['imp_mx']  = get(row,m,'mx')
                fdata[m]['imp_dls'] = get(row,m,'dls')
        elif label == 'AMEX':
            for m in MESES:
                fdata[m]['amex_mx']  = get(row,m,'mx')
                fdata[m]['amex_dls'] = get(row,m,'dls')
        elif label == 'COSTOS FIJOS':
            for m in MESES:
                fdata[m]['cf_mx']  = get(row,m,'mx')
                fdata[m]['cf_dls'] = get(row,m,'dls')

    wb.close()
    return fdata, costos_horas

# ── 4. HORAS detalle (hoja 2026) ─────────────────────────────────────────────
def extraer_horas():
    print("  Horas detalle (hoja 2026)...")
    wb = openpyxl.load_workbook(os.path.abspath(PAGOS_FILE), read_only=True, data_only=True)
    ws = wb['2026']
    data   = {m:{'hours':0.0,'st':0.0,'ot':0.0,'dt':0.0,'res':set()} for m in MESES}
    bd_tmp = {m:{} for m in MESES}

    for row in ws.iter_rows(min_row=2, values_only=True):
        product = str(row[3]).strip().upper() if row[3] else ""
        if product != 'HOURS': continue
        wm = row[11]
        if not isinstance(wm, datetime.datetime) or wm.year != 2026: continue
        if not (1 <= wm.month <= 8): continue
        m = MESES[wm.month-1]

        hrs  = sf(row[13]); st = sf(row[14]); ot = sf(row[15]); dt = sf(row[16])
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
        items=[]
        for (_,_,_),v in bd_tmp[m].items():
            t=v['st']+v['ot']+v['dt']
            items.append({'s':v['s'],'po':v['po'],'c':v['c'],
                          'st':round(v['st'],2),'ot':round(v['ot'],2),
                          'dt':round(v['dt'],2),'t':round(t,2),'r':len(v['r'])})
        breakdown[m]=sorted(items,key=lambda x:-x['t'])
    wb.close()
    return data, breakdown

# ── 5. COSTOS mensuales (hoja COSTOS) ────────────────────────────────────────
def extraer_costos():
    print("  Costos mensuales (hoja COSTOS)...")
    wb = openpyxl.load_workbook(os.path.abspath(PAGOS_FILE), read_only=True, data_only=True)
    ws = wb['COSTOS']
    rows = list(ws.iter_rows(values_only=True))
    costos_mes = {}
    for row in rows[13:27]:
        fecha = row[1]
        if isinstance(fecha, datetime.datetime) and fecha.year==2026 and 1<=fecha.month<=8:
            m = MESES[fecha.month-1]
            costos_mes[m]={'horas_usd':sf(row[2]),'costo_hrs_usd':sf(row[3]),
                           'perdiem_usd':sf(row[4]),'total_usd':sf(row[5]),
                           'horas_mxn':sf(row[6]),'costo_mxn':sf(row[7]),'total_mxn':sf(row[8])}
    for m in MESES:
        if m not in costos_mes:
            costos_mes[m]={k:0.0 for k in ['horas_usd','costo_hrs_usd','perdiem_usd',
                                             'total_usd','horas_mxn','costo_mxn','total_mxn']}
    wb.close()
    return costos_mes

# ── ACTUALIZAR HTML ───────────────────────────────────────────────────────────
MARCA_INI = "// ─── AUTO-GENERADO POR regenerar_dashboard_v2.py ─────────────────"
MARCA_FIN = "// ─── FIN DATOS AUTO-GENERADOS ──────────────────────────────"

def _replace_var(html, varname, new_value_str):
    """Reemplaza var NAME=...;  o var NAME = ...;  con nuevo valor."""
    # Encontrar 'var NAME'
    pos = html.find('var ' + varname)
    if pos < 0:
        print(f"  WARN: no se encontró 'var {varname}' en el HTML")
        return html
    # Encontrar el { de apertura y el } de cierre
    i = html.find('{', pos)
    if i < 0: return html
    depth = 0
    while i < len(html):
        if html[i] == '{': depth += 1
        elif html[i] == '}':
            depth -= 1
            if depth == 0:
                end = i + 1
                # Consumir ; si existe
                j = end
                while j < len(html) and html[j] in ' \t': j+=1
                if j < len(html) and html[j] == ';': end = j+1
                break
        i += 1
    else:
        return html
    new_decl = f"var {varname}={new_value_str};"
    html = html[:pos] + new_decl + html[end:]
    print(f"  OK var {varname} reemplazado")
    return html

def actualizar_html(clientes_mes, total_mes, ing_usd, tc_mes,
                    horas_data, breakdown, costos_mes, fdata, costos_horas):
    if not os.path.exists(OUTPUT_HTML):
        print(f"  ERROR: no encontrado:\n    {OUTPUT_HTML}"); sys.exit(1)

    print("  Leyendo dashboard.html...")
    with open(OUTPUT_HTML,'r',encoding='utf-8') as f: html=f.read()

    # ── var DATA ─────────────────────────────────────────────────────────────
    data_js = {m:{'hours':round(horas_data[m]['hours'],0),'st':round(horas_data[m]['st'],0),
                  'ot':round(horas_data[m]['ot'],0),'dt':round(horas_data[m]['dt'],0),
                  'res':len(horas_data[m]['res'])} for m in MESES}
    html = _replace_var(html, 'DATA', json.dumps(data_js))

    # ── var BREAKDOWN ─────────────────────────────────────────────────────────
    html = _replace_var(html, 'BREAKDOWN', json.dumps(breakdown, ensure_ascii=False))

    # ── var FDATA ─────────────────────────────────────────────────────────────
    fdata_rounded = {m:{k:round(v,2) for k,v in row.items()} for m,row in fdata.items()}
    html = _replace_var(html, 'FDATA', json.dumps(fdata_rounded, ensure_ascii=False))

    # ── var COSTOS ────────────────────────────────────────────────────────────
    # Construir COSTOS con estructura mes -> {usd_hrs, usd_total, mxn_hrs, mxn_total, by_cust, tc}
    costos_js = {}
    for m in MESES:
        ch = costos_horas[m]
        cm = costos_mes[m]
        tc  = tc_mes.get(m, 0)
        costos_js[m] = {
            'usd_hrs':   round(cm['horas_usd'],2),
            'usd_total': round(cm['total_usd'],2),
            'mxn_hrs':   round(cm['horas_mxn'],2),
            'mxn_total': round(cm['total_mxn'],2),
            'by_cust':   [],   # sin desglose por cliente (datos no disponibles a ese nivel)
            'tc':        round(tc, 4),
        }
    html = _replace_var(html, 'COSTOS', json.dumps(costos_js, ensure_ascii=False))

    # ── Bloque de marcadores ──────────────────────────────────────────────────
    if MARCA_INI not in html or MARCA_FIN not in html:
        print("  ERROR: marcadores no encontrados"); sys.exit(1)

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
    L.append("var FCOSTOS_MES = "+json.dumps({m:{k:round(v,2) for k,v in costos_mes[m].items()} for m in MESES},ensure_ascii=False)+";")

    bloque = MARCA_INI+"\n"+"\n".join(L)+"\n"+MARCA_FIN
    idx_i = html.index(MARCA_INI)
    idx_f = html.index(MARCA_FIN)+len(MARCA_FIN)
    html  = html[:idx_i]+bloque+html[idx_f:]
    print("  OK bloque marcadores reemplazado")

    with open(OUTPUT_HTML,'w',encoding='utf-8') as f: f.write(html)
    print(f"  OK guardado")

# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    print("="*60)
    print("  MATUK Dashboard - Regenerando desde Excel")
    print("="*60); print()

    for ruta,nombre in [(FLUJO_FILE,"Flujo de Caja 2026.xlsx"),
                        (PAGOS_FILE,"Servicio Administrativo Pagos 2026.xlsx"),
                        (OUTPUT_HTML,"dashboard.html")]:
        abs_ = os.path.abspath(ruta) if not os.path.isabs(ruta) else ruta
        if os.path.exists(abs_): print(f"  OK {nombre}")
        else: print(f"  NO ENCONTRADO: {nombre}\n    {abs_}"); sys.exit(1)

    print()
    clientes_mes, total_mes = extraer_ingresos()
    ing_usd, tc_mes         = extraer_usd_tc()
    fdata, costos_horas     = extraer_flujo_detalle()
    horas_data, breakdown   = extraer_horas()
    costos_mes              = extraer_costos()

    print()
    print("Resumen:")
    for m in MESES:
        d=horas_data[m]
        print(f"  {m}: {round(d['hours'])}h  ING={total_mes[m]:>12,.0f} MXN  / {ing_usd[m]:>8,.0f} USD")
    print()

    actualizar_html(clientes_mes, total_mes, ing_usd, tc_mes,
                    horas_data, breakdown, costos_mes, fdata, costos_horas)

    print()
    print("="*60)
    print("  LISTO - Dashboard actualizado")
    print(f"  Archivo: {OUTPUT_HTML}")
    print("="*60)

if __name__ == "__main__":
    main()
