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

FLUJO_FILE = r"C:\Users\matuk\Matuk Automation service\Hugo Carreon - BANCOS\Flujo de Caja 2026.xlsx"
PAGOS_FILE = r"C:\Users\matuk\Matuk Automation service\Matuk Automation Repository - Documentos\Reporte de horas 2022\Servicio Administrativo Pagos 2026.xlsx"
OUTPUT_HTML = r"C:\Users\matuk\Desktop\MATUK Dashboard 2026\dashboard.html"


print(f"FLUJO : {FLUJO_FILE}")
print(f"PAGOS : {PAGOS_FILE}")
print(f"HTML  : {OUTPUT_HTML}")

MESES = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT']
MES_MAP = {1:'ENE',2:'FEB',3:'MAR',4:'ABR',5:'MAY',6:'JUN',7:'JUL',8:'AGO',9:'SEP',10:'OCT'}
CMONTHS = ['NOV','DIC','ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT']

MES_LARGO = {
    'ENERO':'ENE','FEBRERO':'FEB','MARZO':'MAR','MARZO ':'MAR','ENERO ':'ENE',
    'ABRIL':'ABR','MAYO':'MAY','JUNIO':'JUN','JULIO':'JUL','AGOSTO':'AGO','SEPTIEMBRE':'SEP',
    'FEBRERO ':'FEB','ABRIL ':'ABR','MAYO ':'MAY','JUNIO ':'JUN',
    'JULIO ':'JUL','AGOSTO ':'AGO','SEPTIEMBRE':'SEP','SEPTIEMBRE ':'SEP',
}

EXCLUIR_ING = {
    'COB','COMPENSACION POR RETRASO','TOTAL GENERAL',
    'DEVOLUCION DEPOSITO EN GARANTIA','DEVOLUCION GUILLERMO FORTINO',
    'DEVOLUCION RAFAEL RODRIGUEZ','PAGO PRESTAMO GUILLERMO FORTINO',
    'PAGO PRESTAMO RAFAEL RODRIGUEZ','OTROS INGRESOS',
    'TOTAL COSTOS','COSTOS','TOTAL GASTOS','AMEX',
}
# NOTA: 'TRASPASO' se eliminó de EXCLUIR_ING para que aparezca en Flujo de Caja e ingresos.

def sf(v):
    try: return float(v) if v is not None else 0.0
    except: return 0.0

# ── INGRESOS MXN ─────────────────────────────────────────────────────────────
def extraer_ingresos():
    print("  Ingresos MXN (Tabla )...")
    wb = openpyxl.load_workbook(FLUJO_FILE, read_only=True, data_only=True)
    ws = wb['Tabla ']
    rows = list(ws.iter_rows(values_only=True))
    # Tabla tiene DOS secciones con headers distintos:
    # Seccion 1 (ingresos): header fila 2 (idx 1) — B=ENERO, C=FEB, D=MAR...
    # Seccion 2 (egresos con TRASPASO/EST FI): header fila 165 (idx 164) — C=ENERO, D=FEB...
    # Leer ambas con su propio col_map para asignar correctamente cada mes

    def make_col_map(header_row):
        cm = {}
        for j,h in enumerate(header_row):
            if isinstance(h,str):
                k = h.strip().upper()
                if k in MES_LARGO and MES_LARGO[k] not in cm:
                    cm[MES_LARGO[k]] = j
        return cm

    # Encontrar indices de ambos headers
    sec1_idx = 1   # fila 2: ingresos (COB, MAGNIT, etc.)
    sec2_idx = None
    for i, row in enumerate(rows):
        vals = [str(v).strip().upper() for v in row if v]
        if 'ENERO' in vals and 'FEBRERO' in vals and 'SEPTIEMBRE' in vals:
            if row[0] and 'HORAS' in str(row[0]).upper():
                sec2_idx = i

    col_map1 = make_col_map(rows[sec1_idx])
    col_map2 = make_col_map(rows[sec2_idx]) if sec2_idx else {}

    # No leer nada de la seccion 2 — TRASPASO y VENTA ACTIVO FIJO ya estan en seccion 1 (ingresos)
    # EST FI es egreso y se lee en extraer_fegr_extra()
    EGRESOS_INGRESOS = set()  # nada de la seccion 2

    clientes_mes = {m:{} for m in MESES}
    total_mes    = {m:0.0 for m in MESES}

    # Leer seccion 1 (ingresos) — parar en primer Total general
    for row in rows[sec1_idx+1:]:
        nombre = str(row[0]).strip().upper() if row[0] else ""
        if nombre == 'TOTAL GENERAL':
            break
        if not nombre or nombre in EXCLUIR_ING: continue
        for m,col in col_map1.items():
            if col < len(row):
                val = sf(row[col])
                if val > 0:
                    clientes_mes[m][nombre] = clientes_mes[m].get(nombre,0.0)+val
                    total_mes[m] += val

    # Leer seccion 2 (egresos) — solo tomar TRASPASO, EST FI y VENTA ACTIVO FIJO
    if sec2_idx and col_map2:
        for row in rows[sec2_idx+1:]:
            nombre = str(row[0]).strip().upper() if row[0] else ""
            if nombre == 'TOTAL GENERAL':
                break
            if nombre not in EGRESOS_INGRESOS: continue
            for m,col in col_map2.items():
                if col < len(row):
                    val = sf(row[col])
                    if val > 0:
                        clientes_mes[m][nombre] = clientes_mes[m].get(nombre,0.0)+val
                        total_mes[m] += val

    wb.close()
    return clientes_mes, total_mes

# ── USD + TC ──────────────────────────────────────────────────────────────────
def extraer_usd_tc():
    print("  USD y TC (FLUJO detalle)...")
    wb = openpyxl.load_workbook(FLUJO_FILE, read_only=True, data_only=True)
    ws = wb['FLUJO detalle']
    rows = list(ws.iter_rows(values_only=True))

    # Buscar fila de encabezados de mes (ENERO, FEBRERO, etc.)
    header_row_idx = None
    for i, row in enumerate(rows[:10]):
        for cell in row:
            if isinstance(cell, str) and cell.strip().upper() in MES_LARGO:
                header_row_idx = i
                break
        if header_row_idx is not None:
            break

    mes_col_dls = {}  # mes -> columna DLS
    if header_row_idx is not None:
        header_row = rows[header_row_idx]
        for j, cell in enumerate(header_row):
            if isinstance(cell, str):
                k = cell.strip().upper().rstrip()
                if k in MES_LARGO:
                    m = MES_LARGO[k]
                    if m not in mes_col_dls:
                        mes_col_dls[m] = j - 1  # columna DLS está antes del nombre

    ing_usd = {m:0.0 for m in MESES}
    tc_mes  = {m:0.0 for m in MESES}

    # Leer TC desde fila anterior al encabezado
    if header_row_idx is not None and header_row_idx > 0:
        tc_row = rows[header_row_idx - 1]
        for m, col in mes_col_dls.items():
            for try_col in [col+1, col, col+2]:
                v = sf(tc_row[try_col]) if try_col < len(tc_row) else 0.0
                if 10 < v < 25:
                    tc_mes[m] = v
                    break

    # Leer INGRESOS USD de fila TOTAL INGRESOS, columna DLS
    for row in rows:
        etiq = str(row[0]).strip().upper() if row[0] else ""
        if etiq in ('TOTAL INGRESOS', 'INGRESOS'):
            for m, col in mes_col_dls.items():
                if col < len(row):
                    ing_usd[m] = sf(row[col])
            break

    wb.close()
    print(f"    TC SEP={tc_mes.get('SEP',0)}  ING_USD SEP={ing_usd.get('SEP',0):,.0f}")
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

# ── PASIVO desde hoja HISTORIAL PASIVOS ──────────────────────────────────────
def extraer_pasivo():
    print("  Pasivo (hoja HISTORIAL PASIVOS)...")
    wb = openpyxl.load_workbook(PAGOS_FILE, read_only=True, data_only=True)
    ws = wb['HISTORIAL PASIVOS']
    rows = list(ws.iter_rows(values_only=True))

    # NOTA: datos en columna B (índice 1), no columna A (índice 0) — col A siempre vacía
    # Fila 3:  col1='CIERRE VIVO '
    # Fila 5:  datos cierre vivo (col1=mes, col2=tc, col3=personas, col4=horas, col5=total_usd, col6=base_mxn, ...)
    # Fila 7:  'CIERRES HISTÓRICOS'
    # Fila 9+: datos históricos  (col1=mes, col2=tc, col3=personas, col4=horas, col5=total_usd, col6=total_mxn, col7=neto_mxn, col8=neto_usd)
    # Fila 15: 'DETALLE HISTÓRICO POR CLIENTE'
    # Fila 17+: detalle (col1=mes, col2=cliente, col3=personas, col4=horas, col5=total_usd, col6=total_mxn, col7=neto_mxn, col8=neto_usd)

    MES_STR = {
        'ene-26':'ENE','feb-26':'FEB','mar-26':'MAR','abr-26':'ABR',
        'may-26':'MAY','jun-26':'JUN','jul-26':'JUL','ago-26':'AGO','sep-26':'SEP',
    }

    cierres = {}
    vivo = None
    detalle = {}
    mode = None  # 'vivo', 'historico', 'detalle'

    for row in rows:
        if not row or len(row) < 2: continue
        col1 = str(row[1]).strip() if row[1] is not None else ''
        col1u = col1.upper()

        if 'CIERRE VIVO' in col1u:
            mode = 'vivo'; continue
        if 'CIERRES HIST' in col1u:
            mode = 'historico'; continue
        if 'DETALLE HIST' in col1u:
            mode = 'detalle'; continue
        if 'C' in col1u and 'MO CERRAR' in col1u:
            mode = None; continue

        mes_str = col1.lower()

        if mode == 'vivo' and mes_str in MES_STR:
            vivo = mes_str
            m = MES_STR[mes_str]
            cierres[m] = {
                'tc': sf(row[2]), 'personas': sf(row[3]),
                'horas': sf(row[4]), 'total_usd': sf(row[5]),
                'total_mxn': sf(row[6]),
                'neto_mxn': 0, 'neto_usd': 0,
                'is_vivo': True,
            }

        elif mode == 'historico' and mes_str in MES_STR:
            m = MES_STR[mes_str]
            cierres[m] = {
                'tc': sf(row[2]), 'personas': sf(row[3]),
                'horas': sf(row[4]), 'total_usd': sf(row[5]),
                'total_mxn': sf(row[6]), 'neto_mxn': sf(row[7]),
                'neto_usd': sf(row[8]) if len(row)>8 and row[8] else 0,
                'is_vivo': False,
            }

        elif mode == 'detalle' and mes_str in MES_STR:
            m = MES_STR[mes_str]
            cliente = str(row[2]).strip() if row[2] else ''
            if not cliente or cliente.upper() in ('CLIENTE', 'TOTAL', 'MES CIERRE'): continue
            if 'TOTAL' in cliente.upper(): continue
            if m not in detalle: detalle[m] = []
            detalle[m].append({
                'c': cliente,
                'p': int(sf(row[3])),
                'h': round(sf(row[4]), 2),
                'usd': round(sf(row[5]), 2),
                'mxn': round(sf(row[6]), 2),
                'neto_mxn': round(sf(row[7]), 2),
                'neto_usd': round(sf(row[8]), 2) if len(row)>8 and row[8] else 0,
            })

    wb.close()
    print(f"    Cierres: {list(cierres.keys())}  Vivo: {vivo}")
    for m, d in cierres.items():
        print(f"    {m}: horas={d['horas']:.1f} usd=${d['total_usd']:,.0f} neto_mxn=${d['neto_mxn']:,.0f}")
    return cierres, detalle, vivo

# ── HORAS detalle para DATA y BREAKDOWN ──────────────────────────────────────
def extraer_horas():
    print("  Horas detalle (hoja 2026)...")
    wb = openpyxl.load_workbook(PAGOS_FILE, read_only=True, data_only=True)
    ws = wb['2026']
    data   = {m:{'hours':0.0,'st':0.0,'ot':0.0,'dt':0.0,'res':set()} for m in MESES}
    bd_tmp = {m:{} for m in MESES}
    nm_tmp = {m:{} for m in MESES}  # nombres con horas por (svc,po,cust)
    cres_tmp = {m:{} for m in MESES}  # CRES: keyed por (po,name,svc,cust)

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[3] or 'HOURS' not in str(row[3]).upper(): continue
        wm = row[11]
        if not isinstance(wm, datetime.datetime) or wm.year != 2026: continue
        if wm.month not in MES_MAP: continue
        m = MES_MAP[wm.month]

        hrs=sf(row[13]); st=sf(row[14]); ot=sf(row[15]); dt=sf(row[16])
        name = str(row[2]).strip() if row[2] else ""
        svc  = str(row[4]).strip() if row[4] else ""
        po   = str(row[5]).strip() if row[5] else ""
        cust = str(row[6]).strip().upper() if row[6] else ""
        proj = str(row[7]).strip() if row[7] else ""
        wd   = sf(row[12])  # WORKED DAYS (col 12)
        try: rate_val = float(row[21]) if row[21] else 0.0
        except: rate_val = 0.0
        try: rate_viat = float(row[24]) if len(row) > 24 and row[24] else 0.0
        except: rate_viat = 0.0
        # Calcular costo de horas desde horas del row × rate (NO usar row[28]/row[30] que son YTD acumulado)
        hrs_row  = st + ot + dt  # horas reales de este row
        days_row = wd if wd > 0 else (round(hrs_row / 8) if hrs_row > 0 else 0)
        perdiem_row = rate_viat * days_row  # costo viáticos de este row (USD)

        data[m]['hours']+=hrs; data[m]['st']+=st; data[m]['ot']+=ot; data[m]['dt']+=dt
        if name: data[m]['res'].add(name)

        key = (svc,po,cust)
        if key not in bd_tmp[m]:
            bd_tmp[m][key]={'s':svc,'po':po,'c':cust,'st':0,'ot':0,'dt':0,'r':set()}
        bd_tmp[m][key]['st']+=st; bd_tmp[m][key]['ot']+=ot; bd_tmp[m][key]['dt']+=dt
        if name: bd_tmp[m][key]['r'].add(name)
        if name:
            if key not in nm_tmp[m]: nm_tmp[m][key]={}
            if name not in nm_tmp[m][key]: nm_tmp[m][key][name]={'st':0,'ot':0,'dt':0,'t':0,'proj':''}
            nm_tmp[m][key][name]['st']+=st; nm_tmp[m][key][name]['ot']+=ot
            nm_tmp[m][key][name]['dt']+=dt; nm_tmp[m][key][name]['t']+=hrs_row
            if proj and not nm_tmp[m][key][name]['proj']: nm_tmp[m][key][name]['proj']=proj
        # CRES: separar por divisa (USA→usd_hrs/usd_cost, MXN→mxn_hrs/mxn_cost)
        divisa = str(row[19]).strip().upper() if len(row) > 19 and row[19] else 'USA'
        es_mxn = (divisa == 'MXN')
        if name and svc and po and cust:
            ckey = (po, name, svc, cust)
            if ckey not in cres_tmp[m]:
                cres_tmp[m][ckey] = {'n':name,'po':po,'svc':svc,'c':cust,'proj':proj,
                                     'rate':rate_val,'rate_viat':rate_viat,
                                     'usd_hrs':0.0,'mxn_hrs':0.0,
                                     'usd_cost':0.0,'mxn_cost':0.0,
                                     'perdiem':0.0,'days':0.0,
                                     'st':0.0,'ot':0.0,'dt':0.0}
            if es_mxn:
                cres_tmp[m][ckey]['mxn_hrs']  += hrs_row           # horas reales MXN
                cres_tmp[m][ckey]['mxn_cost'] += hrs_row * rate_val  # costo en MXN
            else:
                cres_tmp[m][ckey]['usd_hrs']  += hrs_row           # horas reales USD
                cres_tmp[m][ckey]['usd_cost'] += hrs_row * rate_val  # costo en USD
            cres_tmp[m][ckey]['perdiem'] += perdiem_row   # costo viáticos (USD)
            cres_tmp[m][ckey]['days']    += days_row      # dias trabajados
            if rate_viat > cres_tmp[m][ckey]['rate_viat']: cres_tmp[m][ckey]['rate_viat'] = rate_viat  # tomar el maximo del mes
            cres_tmp[m][ckey]['st'] += st
            cres_tmp[m][ckey]['ot'] += ot
            cres_tmp[m][ckey]['dt'] += dt
            if proj and not cres_tmp[m][ckey]['proj']: cres_tmp[m][ckey]['proj'] = proj

    breakdown={}
    for m in MESES:
        items=[{'s':v['s'],'po':v['po'],'c':v['c'],
                'st':round(v['st'],2),'ot':round(v['ot'],2),'dt':round(v['dt'],2),
                't':round(v['st']+v['ot']+v['dt'],2),'r':len(v['r'])}
               for v in bd_tmp[m].values()]
        breakdown[m]=sorted(items,key=lambda x:-x['t'])
    # Construir RESOURCES: detalle de nombres por (svc,po,cust) y mes
    resources={}
    for m in MESES:
        res_items=[]
        for key,nms in nm_tmp[m].items():
            svc2,po2,cust2 = key
            names_list=[{'n':nm,'proj':d.get('proj',''),'st':round(d['st'],2),'ot':round(d['ot'],2),
                         'dt':round(d['dt'],2),'t':round(d['t'],2)}
                        for nm,d in sorted(nms.items(),key=lambda x:-x[1]['t'])]
            res_items.append({'s':svc2,'po':po2,'c':cust2,'names':names_list})
        resources[m]=sorted(res_items,key=lambda x:-sum(n['t'] for n in x['names']))
    # Construir CRES: cres_tmp[m] ya es dict keyed por (po,name,svc,cust)
    cres = {}
    for m in MESES:
        entries = sorted(cres_tmp[m].values(),
                         key=lambda x: -((x['usd_cost']+x.get('perdiem',0))+x['mxn_cost']))
        for e in entries:
            e['usd_hrs']  = round(e['usd_hrs'], 2)   # horas reales USD workers
            e['mxn_hrs']  = round(e['mxn_hrs'], 2)   # horas reales MXN workers
            e['usd_cost'] = round(e['usd_cost'], 2)  # costo en USD
            e['mxn_cost'] = round(e['mxn_cost'], 2)  # costo en MXN
            e['perdiem']  = round(e.get('perdiem', 0), 2)  # costo viáticos USD
            e['days']     = int(e.get('days', 0))          # dias trabajados
            e['st']  = round(e['st'], 2)
            e['ot']  = round(e['ot'], 2)
            e['dt']  = round(e['dt'], 2)
        cres[m] = entries
    wb.close()
    return data, breakdown, resources, cres

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
        (2026,10):'OCT',
    }
    for row in rows[9:26]:  # filas con datos mensuales (hasta OCT)
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
            'usd_hrs':    sf(row[3]),   # Costo Hrs USD
            'perdiem_usd':sf(row[4]),   # Perdiem USD  ← NUEVO
            'usd_total':  sf(row[5]),   # Total USD
            'mxn_hrs':    sf(row[7]),   # Costo Hrs MXN
            'mxn_total':  sf(row[8]),   # Total MXN
            'by_cust':    [],
        }

    # 2. by_cust desde hoja 2026 — col28=Costo Hrs TOTAL USD, col30=Costo TOTAL USD
    ws2 = wb['2026']
    by_cust_tmp = {m:{} for m in list(MES_KEY.values()) + list(CMONTHS)}

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
            by_cust_tmp[m][cust] = {'usd_hrs':0,'usd':0,'mxn_hrs':0,'mxn':0,'by_po':{},'by_svc':{}}
        c = by_cust_tmp[m][cust]
        if divisa == 'MXN':
            tc = tc_mes.get(m, 18.5) if m in MESES else 18.5
            c['mxn_hrs'] += costo_hrs_usd
            c['mxn']     += costo_total_usd
        else:
            c['usd_hrs'] += costo_hrs_usd
            c['usd']     += costo_total_usd

        # Acumular by_svc
        if svc:
            if svc not in c['by_svc']:
                c['by_svc'][svc] = {'svc':svc,'usd_hrs':0,'usd':0}
            c['by_svc'][svc]['usd_hrs'] += costo_hrs_usd if divisa!='MXN' else 0
            c['by_svc'][svc]['usd']     += costo_total_usd if divisa!='MXN' else 0

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
            by_svc = [{'svc':sv['svc'],'usd_hrs':round(sv['usd_hrs'],2),'usd':round(sv['usd'],2)}
                      for sv in sorted(cv.get('by_svc',{}).values(),key=lambda x:-x['usd'])]
            bycust.append({'c':cname,
                           'usd_hrs':round(cv['usd_hrs'],2),'usd':round(cv['usd'],2),
                           'mxn_hrs':round(cv['mxn_hrs'],2),'mxn':round(cv['mxn'],2),
                           'by_svc':by_svc,'by_po':by_po})
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
    """Reemplaza TODAS las ocurrencias de 'var VARNAME={...};' en el HTML.
    Usa 'var VARNAME=' (con =) para evitar match parcial con variables de nombre similar
    (p.ej. var COSTOS no debe hacer match en var COSTOS_OTROS_USD).
    """
    replacement = f"var {varname}={new_value_str};"
    # Patron exacto: 'var VARNAME=' seguido de '{' o '[' (primer char del valor)
    needle = f'var {varname}='
    count = 0
    search_from = 0
    while True:
        pos = html.find(needle, search_from)
        if pos < 0:
            break
        # Buscar el primer '{' o '[' despues del '='
        eq_pos = pos + len(needle)
        # Saltar espacios/tabs entre '=' y valor
        i = eq_pos
        while i < len(html) and html[i] in ' \t': i += 1
        if i >= len(html): break
        open_char = html[i]
        if open_char == '{':
            close_char = '}'
        elif open_char == '[':
            close_char = ']'
        else:
            # Valor escalar (string, numero) — buscar ';'
            semi = html.find(';', eq_pos)
            if semi < 0: break
            end = semi + 1
            html = html[:pos] + replacement + html[end:]
            search_from = pos + len(replacement)
            count += 1
            continue
        depth = 0
        j = i
        while j < len(html):
            if html[j] == open_char: depth += 1
            elif html[j] == close_char:
                depth -= 1
                if depth == 0:
                    end = j+1
                    k = end
                    while k < len(html) and html[k] in ' \t': k+=1
                    if k < len(html) and html[k] == ';': end = k+1
                    break
            j += 1
        else:
            break
        html = html[:pos] + replacement + html[end:]
        search_from = pos + len(replacement)
        count += 1
    if count == 0:
        print(f"  WARN: no se encontró 'var {varname}'")
    else:
        print(f"  OK var {varname} ({count} ocurrencia(s))")
    return html

MARCA_INI = "// ─── AUTO-GENERADO POR regenerar_dashboard_v2.py"
MARCA_FIN = "// ─── FIN DATOS AUTO-GENERADOS"

def actualizar_html(clientes_mes, total_mes, ing_usd, tc_mes,
                    horas_data, breakdown, resources, cres, costos, fdata, fegr_extra=None):
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

    # var RESOURCES (detalle de nombres por mes)
    html = _replace_var(html, 'RESOURCES', json.dumps(resources, ensure_ascii=False))

    # var CRES (detalle de recursos con costos por mes)
    html = _replace_var(html, 'CRES', json.dumps(cres, ensure_ascii=False))

    # var RATE_VIAT: construir desde CRES (todos los workers con su rate diario de viáticos)
    # Esto asegura que el badge "$x/día" aparezca para TODOS los workers, incluyendo los de Kolbenschmidt
    _rate_viat_map = {}
    for _m, _entries in cres.items():
        for _e in _entries:
            _name = _e.get('n', '').upper()
            _rv = _e.get('rate_viat', 0)
            if _name and (_name not in _rate_viat_map or _rv > 0):
                _rate_viat_map[_name] = _rv
    html = _replace_var(html, 'RATE_VIAT', json.dumps(_rate_viat_map, ensure_ascii=False))

    # var FDATA
    fdata_r = {m:{k:round(v,2) for k,v in row.items()} for m,row in fdata.items()}
    html = _replace_var(html, 'FDATA', json.dumps(fdata_r, ensure_ascii=False))

    # var COSTOS
    costos_r = {}
    for m in CMONTHS:
        if m in costos:
            d = costos[m]
            costos_r[m] = {
                'usd_hrs':    round(d['usd_hrs'],2),
                'perdiem_usd':round(d.get('perdiem_usd',0),2),
                'usd_total':  round(d['usd_total'],2),
                'mxn_hrs':    round(d['mxn_hrs'],2),
                'mxn_total':  round(d['mxn_total'],2),
                'by_cust':    d['by_cust'],
                'tc':         d.get('tc',0),
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
        # FEGR_EXTRA no va en el bloque — se actualiza via _replace_var abajo

        bloque = MARCA_INI+"\n"+"\n".join(L)+"\n"+MARCA_FIN
        html  = html[:idx_i]+bloque+html[idx_f_end:]
        print("  OK bloque marcadores")
    else:
        print("  WARN: marcadores no encontrados")

    # Actualizar FEGR_EXTRA via _replace_var (sobrescribe la vieja hardcoded)
    if fegr_extra:
        html = _replace_var(html, 'FEGR_EXTRA', json.dumps(fegr_extra, ensure_ascii=False))


    with open(OUTPUT_HTML,'w',encoding='utf-8') as f: f.write(html)
    print(f"  OK guardado ({len(html):,} bytes)")



# ── FEGR_EXTRA desde FLUJO detalle ───────────────────────────────────────────
def extraer_fegr_extra():
    """Extrae comisiones, gastos_op, seguros, ptu, prestamo, estfi, costos_otros por mes."""
    print("  FEGR_EXTRA (FLUJO detalle)...")
    wb = openpyxl.load_workbook(FLUJO_FILE, read_only=True, data_only=True)
    ws = wb['FLUJO detalle']
    rows = list(ws.iter_rows(values_only=True))

    # Columnas MXN por mes: ENE=col2, FEB=col5, MAR=col8... (0-indexed, cada 3 cols)
    mxn_cols = {m: 2 + i*3 for i, m in enumerate(MESES)}

    def get_mx(row, m):
        col = mxn_cols.get(m, -1)
        if col < 0 or col >= len(row): return 0.0
        return sf(row[col])

    result = {m: {'comisiones':0,'gastos_op':0,'nom_admon':0,'seguros':0,'ptu':0,'prestamo':0,'estfi':0,
                  'activo_fijo':0,'costos_otros':0,'total_costos':0,'horas':0,'viat':0} for m in MESES}

    # Labels after .strip() (trailing spaces removed)
    # NOTA: gastos_op = Total GASTOS - COMISIONES - NOM ADMON (para no duplicar con fdata)
    TARGETS = {
        'COMISIONES':       'comisiones',
        'NOM ADMON':        'nom_admon',   # se resta de gastos_op al final
        'TOTAL GASTOS':     'gastos_op',
        'Total GASTOS':     'gastos_op',
        'SEGUROS':          'seguros',
        'PTU':              'ptu',
        'ESTRATEGIA FISCAL':'estfi',
        'EST FI':           'estfi',
        'ACTIVO FIJO':      'activo_fijo',
        'Total COSTOS':     'total_costos',
        'TOTAL COSTOS':     'total_costos',
        'HORAS':            'horas',
        'VIATICOS':         'viat',
    }
    # PRESTAMO: hay 2 filas — fila 121 (en COSTOS, siempre 0) y fila 144 (en GASTOS, tiene valor)
    # Sumar ambas por si acaso (la primera siempre es 0)
    found = set()

    for i, row in enumerate(rows):
        label = str(row[0]).strip() if row[0] else ""
        # PRESTAMO: sumar todas las ocurrencias (fila 121=0, fila 144=valor real)
        if label == 'PRESTAMO':
            for m in MESES:
                result[m]['prestamo'] += get_mx(row, m)
            continue
        if label in TARGETS and label not in found:
            key = TARGETS[label]
            found.add(label)
            for m in MESES:
                result[m][key] = get_mx(row, m)

    # gastos_op = Total GASTOS - COMISIONES - NOM ADMON - PTU (ya aparecen por separado en fdata/fegr_extra)
    for m in MESES:
        result[m]['gastos_op'] = max(0, result[m]['gastos_op'] - result[m]['comisiones'] - result[m]['nom_admon'] - result[m]['ptu'])

    # costos_otros = Total COSTOS - horas - viaticos
    for m in MESES:
        tc = result[m]['total_costos']
        h  = result[m]['horas']
        v  = result[m]['viat']
        result[m]['costos_otros'] = max(0, round(tc - h - v))

    wb.close()

    # Limpiar campos internos antes de devolver
    final = {}
    for m in MESES:
        d = result[m]
        if any(d[k] for k in ['comisiones','gastos_op','seguros','ptu','prestamo','estfi','costos_otros','activo_fijo']):
            final[m] = {k: int(round(d[k])) for k in ['costos_otros','gastos_op','seguros','comisiones','ptu','prestamo','estfi','activo_fijo']}

    print(f"    Meses con FEGR_EXTRA: {list(final.keys())}")
    if 'SEP' in final:
        print(f"    SEP: {final['SEP']}")
    return final


# ── FIXES POST-REGENERACION ───────────────────────────────────────────────────

# ── SALDO ACTUAL DE BANCOS ────────────────────────────────────────────────────
def extraer_saldo_bancos():
    """Lee el último saldo de los 5 bancos desde la hoja CONCENTRADO de Flujo de Caja 2026.xlsx.
    Retorna MXN como valor principal y USD como sub-valor para cuentas en dólares.
    banks = {'BBVA Pesos': {mxn:...}, 'BBVA USD': {mxn:..., usd:...}, ...}
    """
    import datetime as _dt
    print("  Saldo bancos (CONCENTRADO - 5 bancos)...")
    try:
        wb = openpyxl.load_workbook(FLUJO_FILE, read_only=True, data_only=True)
    except Exception as e:
        print(f"  WARN saldo_bancos: {e}")
        return None

    # Etiquetas en col 12 (idx 11) de CONCENTRADO (con strip para eliminar espacios)
    BANCO_LABELS = {
        'Vs BBVA pesos': {'label': 'BBVA Pesos',  'isUSD': False},
        'Vs BBVA USD':   {'label': 'BBVA USD',    'isUSD': True},
        'BMX USD':       {'label': 'BMX USD',     'isUSD': True},
        'BMXQRO MN':     {'label': 'BMX QRO MN',  'isUSD': False},
        'BANAMEX MN':    {'label': 'Banamex MN',  'isUSD': False},
    }

    try:
        ws_conc = wb['CONCENTRADO']
    except Exception as e:
        print(f"  WARN: hoja CONCENTRADO no encontrada: {e}")
        return None

    rows = list(ws_conc.iter_rows(values_only=True))

    last_found = {}  # label_raw -> {mxn, usd?}
    total_mxn = None
    meses = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic']

    for r in rows:
        if len(r) <= 12:
            continue
        label_cell = r[11]
        if label_cell is None:
            continue
        label = str(label_cell).strip()  # strip: etiquetas pueden tener espacios

        if label in BANCO_LABELS:
            mxn_val = r[12] if len(r) > 12 else None
            usd_val = r[13] if len(r) > 13 else None
            if isinstance(mxn_val, (int, float)):
                entry = {'mxn': round(float(mxn_val), 2)}
                if BANCO_LABELS[label]['isUSD'] and isinstance(usd_val, (int, float)):
                    entry['usd'] = round(float(usd_val), 2)
                last_found[label] = entry
        elif label == 'SALDO':
            mxn_val = r[12] if len(r) > 12 else None
            if isinstance(mxn_val, (int, float)):
                total_mxn = round(float(mxn_val), 2)

    # Fecha del último bloque: en CONCENTRADO la fecha está en col 3 (idx 3)
    # Buscar hacia arriba desde el último banco encontrado
    ultima_fecha_datos = None
    for i in range(len(rows)-1, -1, -1):
        r = rows[i]
        if len(r) > 11 and r[11] is not None:
            label = str(r[11]).strip()
            if label in BANCO_LABELS:
                for j in range(i, max(i-30, -1), -1):
                    rr = rows[j]
                    if rr and len(rr) > 3 and isinstance(rr[3], _dt.datetime):
                        ultima_fecha_datos = rr[3]
                        break
                break
    # Fallback: última fecha en col 3
    if ultima_fecha_datos is None:
        for r in reversed(rows):
            if r and len(r) > 3 and isinstance(r[3], _dt.datetime):
                ultima_fecha_datos = r[3]
                break

    # Construir banks: {label_display: {mxn, usd?}}
    banks = {}
    for label_raw, info in BANCO_LABELS.items():
        if label_raw in last_found:
            banks[info['label']] = last_found[label_raw]

    if total_mxn is None:
        total_mxn = round(sum(v['mxn'] for v in banks.values()), 2)

    fecha_str = None
    if ultima_fecha_datos:
        fecha_str = f"{ultima_fecha_datos.day}-{meses[ultima_fecha_datos.month-1]}-{ultima_fecha_datos.year}"

    import datetime as _dt2
    hoy = _dt2.date.today()
    fecha_gen = f"{hoy.day}-{meses[hoy.month-1]}-{hoy.year}"

    print(f"    Total: {total_mxn:,.0f} MXN | Último dato: {fecha_str} | Generado: {fecha_gen}")
    for lbl, v in banks.items():
        usd_str = f" / USD {v['usd']:,.2f}" if 'usd' in v else ''
        print(f"      {lbl}: MXN {v['mxn']:,.2f}{usd_str}")

    return {'fecha': fecha_str, 'fecha_generacion': fecha_gen, 'total': total_mxn, 'banks': banks}

def aplicar_fixes(html, ultimo_mes_real=None, fegr_extra=None, tc_mes=None):
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
  document.querySelectorAll('.cfilter-btn').forEach(b => { const bm = b.dataset.m || b.getAttribute('onclick').match(/cFilterMes.'([^']+)'/)?.[1] || ''; b.classList.toggle('active', bm === m); });
  // Clear search if active
  const inp = document.getElementById('cSearchInput');
  if(inp && inp.value){ inp.value=''; searchCostosTable(''); }
  buildCostosTable();
}"""
    NEW3 = """function cFilterMes(m){
  cCurrentMes = m;
  document.querySelectorAll('.cfilter-btn').forEach(b => { const bm = b.dataset.m || b.getAttribute('onclick').match(/cFilterMes.'([^']+)'/)?.[1] || ''; b.classList.toggle('active', bm === m); });
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

    # FIX 5: Canvas HTML para graficas costos — deben estar ANTES de la tabla (ctbody)
    CANVAS5 = '\n      <!-- GRAFICAS COSTOS -->\n      <div style="display:flex;gap:20px;padding:0 32px 32px 32px;box-sizing:border-box">\n        <div style="flex:1;background:#fff;border-radius:12px;padding:20px;box-shadow:0 1px 6px rgba(0,0,0,.07)">\n          <div style="font-size:.72rem;font-weight:800;color:#1F3A5F;text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px">\n            <span style="color:#F59E0B;margin-right:6px">&#9679;</span>Tendencia Mensual vs Acumulado\n          </div>\n          <canvas id="costos-tendencia-chart" height="110"></canvas>\n        </div>\n        <div style="flex:1;background:#fff;border-radius:12px;padding:20px;box-shadow:0 1px 6px rgba(0,0,0,.07)">\n          <div style="font-size:.72rem;font-weight:800;color:#1F3A5F;text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px">\n            <span style="color:#1F5BA6;margin-right:6px">&#9679;</span>Total de Costos por Mes (USD)\n          </div>\n          <canvas id="costos-barras-chart" height="110"></canvas>\n        </div>\n      </div>\n'
    TABLE_ANCHOR5 = '<tbody id="ctbody">'
    import re as _re5
    canvas_present = 'costos-tendencia-chart' in html
    # Placeholder comentario sin el div real
    _CANVAS_PLACEHOLDER = '<!-- GRAFICAS COSTOS -->'
    _placeholder_only = _CANVAS_PLACEHOLDER in html and not canvas_present
    canvas_pos5 = html.find('costos-tendencia-chart')
    table_pos5 = html.find(TABLE_ANCHOR5)
    if _placeholder_only:
        # Reemplazar el placeholder por el canvas completo directamente
        html = html.replace(_CANVAS_PLACEHOLDER, CANVAS5.strip(), 1)
        fixes_ok.append("Canvas graficas costos: placeholder reemplazado por canvas real")
    elif not canvas_present:
        # Insertar antes de la tabla
        if TABLE_ANCHOR5 in html:
            # buscar el <table> más cercano antes de ctbody
            t_open = html.rfind('<table>', 0, table_pos5)
            if t_open > 0:
                html = html[:t_open] + CANVAS5 + html[t_open:]
                fixes_ok.append("Canvas graficas costos insertado antes de tabla")
            else:
                fixes_fail.append("Canvas graficas costos (no se encontro <table> antes de ctbody)")
        else:
            fixes_fail.append("Canvas graficas costos (ctbody no encontrado)")
    elif canvas_pos5 > table_pos5:
        # Gráficas están DESPUÉS de la tabla — extraerlas y moverlas antes
        flex_start = html.rfind('<div style="display:flex', 0, canvas_pos5)
        if flex_start == -1:
            # El placeholder comentario está después de tabla pero sin div real — reemplazarlo
            ph_pos = html.rfind(_CANVAS_PLACEHOLDER, 0, canvas_pos5+50)
            if ph_pos > 0:
                html = html[:ph_pos] + html[ph_pos+len(_CANVAS_PLACEHOLDER):]
                table_pos5b = html.find(TABLE_ANCHOR5)
                t_open = html.rfind('<table>', 0, table_pos5b)
                if t_open > 0:
                    html = html[:t_open] + CANVAS5 + html[t_open:]
                    fixes_ok.append("Canvas graficas costos movido antes de tabla (placeholder)")
                else:
                    fixes_fail.append("Canvas graficas costos (movido placeholder pero no se encontro tabla)")
            else:
                fixes_fail.append("Canvas graficas costos (canvas despues de tabla, no hay flex ni placeholder)")
        else:
            # El bloque termina después del canvas de barras
            barras_pos = html.find('costos-barras-chart', canvas_pos5)
            flex_end = html.find('</div>', barras_pos)
            flex_end = html.find('</div>', flex_end+1)  # cierra el inner div
            flex_end = html.find('</div>', flex_end+1)  # cierra el outer flex div
            flex_end += len('</div>')
            if flex_start > 0 and flex_end > flex_start:
                canvas_block = html[flex_start:flex_end]
                # Verificar ANTES de eliminar que la tabla destino existe
                table_pos5b = html.find(TABLE_ANCHOR5)
                t_open = html.rfind('<table>', 0, table_pos5b) if table_pos5b > 0 else -1
                if t_open > 0:
                    # Solo mover si hay destino valido
                    html = html[:flex_start] + html[flex_end:]
                    # Recalcular posicion de la tabla despues de eliminar el bloque
                    table_pos5b2 = html.find(TABLE_ANCHOR5)
                    t_open2 = html.rfind('<table>', 0, table_pos5b2) if table_pos5b2 > 0 else -1
                    if t_open2 > 0:
                        html = html[:t_open2] + '\n' + canvas_block + '\n\n' + html[t_open2:]
                        fixes_ok.append("Canvas graficas costos movido antes de tabla")
                    else:
                        # Tabla desaparecio al mover — reinsertar en posicion original
                        html = html[:flex_start] + canvas_block + html[flex_start:]
                        fixes_fail.append("Canvas graficas costos (tabla desaparecio, canvas restaurado en posicion original)")
                else:
                    # No hay tabla destino — NO mover para no perder los canvas
                    fixes_ok.append("Canvas graficas costos (ya OK — no se mueve, tabla destino no encontrada)")
            else:
                fixes_fail.append("Canvas graficas costos (no se pudo extraer bloque flex)")
    else:
        fixes_ok.append("Canvas graficas costos (ya OK — antes de tabla)")

    # FIX MONTHS / FMONTHS: actualizar arrays de meses segun datos reales de horas
    # var MONTHS usada por pestaña Horas, var FMONTHS por Flujo
    # MONTHS se basa en horas_data (puede tener OCT aunque Flujo no tenga OCT aún)
    # FMONTHS se basa en ultimo_mes_real (flujo)
    import re as _re_months
    _MESES_FULL = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC']
    # Ultimo mes con horas — buscar en var DATA del HTML ya actualizado
    import re as _re_data_h
    import json as _json_h
    _data_match = _re_data_h.search(r'var DATA=({[^;]+});', html)
    _ultimo_horas = 'SEP'
    if _data_match:
        try:
            _data_obj = _json_h.loads(_data_match.group(1))
            for _mh in reversed(_MESES_FULL):
                if _data_obj.get(_mh, {}).get('hours', 0) > 0:
                    _ultimo_horas = _mh
                    break
        except Exception:
            pass
    _idx_horas = _MESES_FULL.index(_ultimo_horas)
    _months_horas_js = json.dumps(_MESES_FULL[:_idx_horas + 1])
    # Último mes con datos de flujo
    _ultimo_flujo = ultimo_mes_real if (ultimo_mes_real and ultimo_mes_real in _MESES_FULL) else 'SEP'
    _idx_flujo = _MESES_FULL.index(_ultimo_flujo)
    _months_flujo_js = json.dumps(_MESES_FULL[:_idx_flujo + 1])
    # Reemplazar var MONTHS (horas) — hasta último mes con horas
    _m = _re_months.search(r'var MONTHS\s*=\s*\[[^\]]*\]\s*;', html)
    if _m:
        _new = f'var MONTHS={_months_horas_js};'
        if html[_m.start():_m.end()] != _new:
            html = html[:_m.start()] + _new + html[_m.end():]
            fixes_ok.append(f'var MONTHS actualizado hasta {_ultimo_horas}')
        else:
            fixes_ok.append(f'var MONTHS ya OK ({_ultimo_horas})')
    else:
        fixes_fail.append('var MONTHS no encontrado')
    # Reemplazar var FMONTHS (flujo) — hasta último mes de flujo
    _fm = _re_months.search(r"var FMONTHS\s*=\s*\[[^\]]*\]\s*;", html)
    if _fm:
        _new = f'var FMONTHS={_months_flujo_js};'
        if html[_fm.start():_fm.end()] != _new:
            html = html[:_fm.start()] + _new + html[_fm.end():]
            fixes_ok.append(f'var FMONTHS actualizado hasta {_ultimo_flujo}')
        else:
            fixes_ok.append(f'var FMONTHS ya OK ({_ultimo_flujo})')
    else:
        fixes_fail.append('var FMONTHS no encontrado')

    # FIX BADGE PERIODO: "ENE – OCT 2026" según último mes con horas
    import re as _re_badge
    _badge_pat = _re_badge.search(r'<span class="badge-period">ENE\s*[–\-]\s*\w+\s*2026</span>', html)
    if _badge_pat:
        _new_badge = f'<span class="badge-period">ENE – {_ultimo_horas} 2026</span>'
        if html[_badge_pat.start():_badge_pat.end()] != _new_badge:
            html = html[:_badge_pat.start()] + _new_badge + html[_badge_pat.end():]
            fixes_ok.append(f'Badge periodo actualizado hasta {_ultimo_horas}')
        else:
            fixes_ok.append(f'Badge periodo ya OK ({_ultimo_horas})')
    else:
        fixes_fail.append('Badge periodo no encontrado')

    # FIX CFILTER BUTTONS: regenerar botones FILTRAR POR MES con meses hasta _ultimo_horas
    _btn_style = 'padding:4px 12px;border-radius:20px;border:1.5px solid #D1D5DB;background:#fff;font-size:.72rem;cursor:pointer;font-weight:600;color:#6B7A99'
    _btn_todos = '<button onclick="cFilterMes(\'TODOS\')" data-m="TODOS" id="cbtn-todos" class="cfilter-btn active" style="padding:4px 12px;border-radius:20px;border:none;background:var(--blue);color:#fff;font-size:.72rem;cursor:pointer;font-weight:600">Todos</button>'
    _btns_meses = '\n      '.join(
        f'<button onclick="cFilterMes(\'{m}\')" data-m="{m}" class="cfilter-btn" style="{_btn_style}">{m}</button>'
        for m in _MESES_FULL[:_idx_horas + 1]
    )
    _new_btn_block = (
        '<span style="font-size:.65rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#6B7A99;margin-right:4px">FILTRAR POR MES:</span>\n'
        f'      {_btn_todos}\n'
        f'      {_btns_meses}'
    )
    _cfilter_ini = html.find('<!-- FILTRAR POR MES -->')
    if _cfilter_ini >= 0:
        _inner_div_start = html.find('<div style="display:flex;gap:6px', _cfilter_ini)
        if _inner_div_start >= 0:
            _inner_div_end = html.find('</div>', _inner_div_start) + len('</div>')
            _old_inner = html[_inner_div_start:_inner_div_end]
            _new_inner = f'<div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center">\n      {_new_btn_block}\n    </div>'
            if _old_inner != _new_inner:
                html = html[:_inner_div_start] + _new_inner + html[_inner_div_end:]
                fixes_ok.append(f'Botones cFilterMes regenerados hasta {_ultimo_horas}')
            else:
                fixes_ok.append(f'Botones cFilterMes ya OK ({_ultimo_horas})')
        else:
            fixes_fail.append('cFilterMes: inner div no encontrado')
    else:
        fixes_fail.append('cFilterMes: marcador <!-- FILTRAR POR MES --> no encontrado')

    # FIX FILTER-BTN (Horas por Mes): regenerar botones filter() hasta _ultimo_horas
    _dq = '"'
    _sq = "'"
    _horas_btn_todos = (
        '<button class=' + _dq + 'filter-btn active' + _dq +
        ' data-m=' + _dq + 'TODOS' + _dq +
        ' onclick=' + _dq + 'filter(' + _sq + 'TODOS' + _sq + ')' + _dq +
        '>Todos</button>'
    )
    def _mk_fhbtn(m):
        return (
            '<button class=' + _dq + 'filter-btn' + _dq +
            ' data-m=' + _dq + m + _dq +
            ' onclick=' + _dq + 'filter(' + _sq + m + _sq + ')' + _dq +
            '>' + m + '</button>'
        )
    _horas_btns = '\n    '.join(_mk_fhbtn(m) for m in _MESES_FULL[:_idx_horas + 1])
    _horas_label = '<span class=' + _dq + 'filter-label' + _dq + '>Filtrar por mes:</span>'
    _horas_label_pos = html.find(_horas_label)
    if _horas_label_pos >= 0:
        _horas_div_end = html.find('</div>', _horas_label_pos)
        if _horas_div_end >= 0:
            _old_horas = html[_horas_label_pos:_horas_div_end]
            _new_horas = _horas_label + '\n    ' + _horas_btn_todos + '\n    ' + _horas_btns
            if _old_horas != _new_horas:
                html = html[:_horas_label_pos] + _new_horas + html[_horas_div_end:]
                fixes_ok.append('Botones filter() Horas regenerados hasta ' + _ultimo_horas)
            else:
                fixes_ok.append('Botones filter() Horas ya OK (' + _ultimo_horas + ')')
        else:
            fixes_fail.append('filter() Horas: div end no encontrado')
    else:
        fixes_fail.append('filter() Horas: span filter-label no encontrado')

    # FIX FLUJO: botones de mes, textos ENE-XXX y subtítulo Saldo Final
    # Determinar último mes con datos en FDATA
    MESES_ORD = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC']
    MES_NOMBRE = {'ENE':'enero','FEB':'febrero','MAR':'marzo','ABR':'abril','MAY':'mayo',
                  'JUN':'junio','JUL':'julio','AGO':'agosto','SEP':'septiembre',
                  'OCT':'octubre','NOV':'noviembre','DIC':'diciembre'}
    import re as _re2
    import json as _json
    MESES_ORD_LOCAL = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC']
    # Usar mes real pasado como parámetro, si existe
    if ultimo_mes_real and ultimo_mes_real in MESES_ORD_LOCAL:
        ultimo_mes_flujo = ultimo_mes_real
    else:
        # Buscar último mes con datos reales (ing_mx > 0) en FDATA
        fdata_m = _re2.search(r'var FDATA=({[\s\S]*?});\s*\n', html)
        ultimo_mes_flujo = 'AGO'
        if fdata_m:
            try:
                fdata_obj = _json.loads(fdata_m.group(1))
                for m in reversed(MESES_ORD_LOCAL):
                    if m in fdata_obj and (fdata_obj[m].get('ing_mx',0) or 0) > 0:
                        ultimo_mes_flujo = m
                        break
            except Exception:
                meses_en_fdata = _re2.findall(r'"([A-Z]{3})":\s*\{', html[fdata_m.start():fdata_m.start()+8000])
                for m in reversed(MESES_ORD_LOCAL):
                    if m in meses_en_fdata:
                        ultimo_mes_flujo = m
                        break

    # Agregar botón del último mes si no existe
    btn_id_ultimo = 'data-m="' + ultimo_mes_flujo + '"'
    if btn_id_ultimo not in html:
        old_ago_btn = "<button class=\"fmes-btn\" data-m=\"AGO\" onclick=\"fSetMes('AGO',this)\">AGO</button>"
        new_buttons = old_ago_btn + '\n    <button class="fmes-btn" data-m="' + ultimo_mes_flujo + '" onclick="fSetMes(\'' + ultimo_mes_flujo + '\',this)">' + ultimo_mes_flujo + '</button>'
        if old_ago_btn in html:
            html = html.replace(old_ago_btn, new_buttons, 1)
            fixes_ok.append("Boton " + ultimo_mes_flujo + " flujo")

    # Actualizar textos ENE–??? al rango correcto
    for old_rng in ['ENE–ENE','ENE–FEB','ENE–MAR','ENE–ABR','ENE–MAY','ENE–JUN',
                    'ENE–JUL','ENE–AGO','ENE–SEP','ENE–OCT','ENE–NOV','ENE-AGO',
                    'ENE-SEP','ENE-ENE','ENE-FEB','ENE-MAR']:
        if old_rng in html:
            html = html.replace(old_rng, f'ENE–{ultimo_mes_flujo}')
            fixes_ok.append(f"Rango flujo {old_rng}→ENE–{ultimo_mes_flujo}")

    # Actualizar subtítulo Saldo Final
    mes_nombre = MES_NOMBRE.get(ultimo_mes_flujo, ultimo_mes_flujo.lower())
    for m_nombre in MES_NOMBRE.values():
        old_sub = f"'Saldo final {m_nombre} 2026'"
        if old_sub in html:
            new_sub = f"'Saldo final {mes_nombre} 2026'"
            html = html.replace(old_sub, new_sub, 1)
            fixes_ok.append(f"Subtítulo saldo final →{mes_nombre}")
            break

    # FIX 6: buildCostosKPIs — respetar cCurrentMes al filtrar por mes
    OLD6a = 'function buildCostosKPIs(){\n  var months=Object.keys(COSTOS);'
    NEW6a = '''function buildCostosKPIs(){
  var allMonths=CMONTHS.filter(m=>COSTOS[m]);
  var months=(typeof cCurrentMes!=="undefined"&&cCurrentMes!=="TODOS")
    ? [cCurrentMes].filter(m=>COSTOS[m]) : allMonths;'''
    if OLD6a in html:
        html = html.replace(OLD6a, NEW6a, 1); fixes_ok.append("buildCostosKPIs filtro mes")
    elif 'allMonths=CMONTHS.filter' in html:
        fixes_ok.append("buildCostosKPIs filtro mes (ya OK)")
    else:
        fixes_fail.append("buildCostosKPIs filtro mes")

    # FIX 6b: cFilterMes llama buildCostosKPIs
    OLD6b = '  buildCostosTable();\n  buildCostosCharts();\n}'
    NEW6b = '  buildCostosTable();\n  buildCostosKPIs();\n  buildCostosCharts();\n}'
    if OLD6b in html and 'buildCostosKPIs' not in html[html.find('function cFilterMes'):html.find('function cFilterMes')+300]:
        html = html.replace(OLD6b, NEW6b, 1); fixes_ok.append("cFilterMes→buildCostosKPIs")
    elif 'buildCostosKPIs' in html[html.find('function cFilterMes'):html.find('function cFilterMes')+300]:
        fixes_ok.append("cFilterMes→buildCostosKPIs (ya OK)")

    # FIX: Título "Indicadores Clave — ENE a MES 2026" dinámico
    for old_mes in ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC']:
        if old_mes == ultimo_mes_flujo:
            continue
        old_t = 'Indicadores Clave — ENE a ' + old_mes + ' 2026'
        new_t = 'Indicadores Clave — ENE a ' + ultimo_mes_flujo + ' 2026'
        if old_t in html:
            html = html.replace(old_t, new_t)
            fixes_ok.append('KPI title ' + old_mes + '->' + ultimo_mes_flujo)

    # Fix perdiem_usd: usar valor real del resumen en lugar de calcular por diferencia
    old_viat = "    const viaticos = d.usd_total - d.usd_hrs - otrosCostos;"
    new_viat = "    const viaticos = (d.perdiem_usd != null && d.perdiem_usd > 0) ? d.perdiem_usd : (d.usd_total - d.usd_hrs - otrosCostos);"
    if old_viat in html:
        html = html.replace(old_viat, new_viat)
        fixes_ok.append('perdiem_usd formula')
    elif new_viat in html:
        fixes_ok.append('perdiem_usd formula(ya)')

    # Fix rsRender allMonths: incluir SEP en el Resumen
    old_rs = "  const allMonths = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO'];"
    new_rs = "  const allMonths = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP'];"
    if old_rs in html:
        html = html.replace(old_rs, new_rs)
        fixes_ok.append('rsRender allMonths +SEP')
    elif new_rs in html:
        fixes_ok.append('rsRender allMonths +SEP (ya OK)')

    # Fix polyfill roundRect para Chrome < 99
    polyfill_marker = 'CanvasRenderingContext2D.prototype.roundRect'
    polyfill_code = '''
// Polyfill para ctx.roundRect (no disponible en browsers < Chrome 99)
if(typeof CanvasRenderingContext2D !== 'undefined' && !CanvasRenderingContext2D.prototype.roundRect){
  CanvasRenderingContext2D.prototype.roundRect = function(x,y,w,h,r){
    var radius = typeof r === 'number' ? r : (Array.isArray(r) ? r[0] : 0);
    this.moveTo(x+radius,y);
    this.lineTo(x+w-radius,y);
    this.quadraticCurveTo(x+w,y,x+w,y+radius);
    this.lineTo(x+w,y+h-radius);
    this.quadraticCurveTo(x+w,y+h,x+w-radius,y+h);
    this.lineTo(x+radius,y+h);
    this.quadraticCurveTo(x,y+h,x,y+h-radius);
    this.lineTo(x,y+radius);
    this.quadraticCurveTo(x,y,x+radius,y);
    this.closePath();
  };
}
'''
    target_fn = 'function fRenderBarIngEgr(){'
    if polyfill_marker not in html and target_fn in html:
        html = html.replace(target_fn, polyfill_code + target_fn, 1)
        fixes_ok.append('polyfill roundRect')
    elif polyfill_marker in html:
        fixes_ok.append('polyfill roundRect (ya OK)')

    # FIX RS MONTH FILTER: agregar botón SEP en Resumen hero section
    OLD_RS_BTN = """      <button class="rs-mes-btn" onclick="setRsMes('AGO')" id="rsmb-AGO">AGO</button>
    </div>
  </div>
  <div class="hero-right">"""
    NEW_RS_BTN = """      <button class="rs-mes-btn" onclick="setRsMes('AGO')" id="rsmb-AGO">AGO</button>
      <button class="rs-mes-btn" onclick="setRsMes('SEP')" id="rsmb-SEP">SEP</button>
    </div>
  </div>
  <div class="hero-right">"""
    if OLD_RS_BTN in html:
        html = html.replace(OLD_RS_BTN, NEW_RS_BTN, 1)
        fixes_ok.append("RS MONTH FILTER botón SEP")
    elif 'rsmb-SEP' in html:
        fixes_ok.append("RS MONTH FILTER botón SEP (ya OK)")
    else:
        fixes_fail.append("RS MONTH FILTER botón SEP")

    # FIX COSTOS_OTROS_USD: declarar variable antes de que buildCostosTable la use
    # Calcular COSTOS_OTROS_USD = costos_otros MXN / TC por mes
    _cou = {}
    if fegr_extra and tc_mes:
        for _m in MESES:
            _co = fegr_extra.get(_m, {}).get('costos_otros', 0)
            _tc = tc_mes.get(_m, 18.5)
            if _co and _tc:
                _cou[_m] = round(_co / _tc, 2)
    _cou_js = 'var COSTOS_OTROS_USD=' + json.dumps(_cou) + ';'
    if 'var COSTOS_OTROS_USD' not in html:
        old_cres = 'var CRES='
        if old_cres in html:
            html = html.replace(old_cres, _cou_js + '\n' + old_cres, 1)
            fixes_ok.append("COSTOS_OTROS_USD declarado con datos")
        else:
            fixes_fail.append("COSTOS_OTROS_USD (anchor CRES no encontrado)")
    else:
        html = _replace_var(html, 'COSTOS_OTROS_USD', json.dumps(_cou))
        fixes_ok.append("COSTOS_OTROS_USD actualizado")

    # FIX COSTOS_TOTAL_USD: sumar otrosCostos al totalTotalUSD (acumulador)
    OLD_TOTAL = "    totalTotalUSD += d.usd_total||0;"
    NEW_TOTAL = "    totalTotalUSD += (d.usd_total||0) + otrosCostos;"
    if OLD_TOTAL in html:
        html = html.replace(OLD_TOTAL, NEW_TOTAL, 1)
        fixes_ok.append("totalTotalUSD incluye otrosCostos")
    elif NEW_TOTAL in html:
        fixes_ok.append("totalTotalUSD (ya OK)")
    else:
        fixes_fail.append("totalTotalUSD (patron no encontrado)")

    # FIX COSTOS_ROW_USD: cada fila muestra usd_total + otrosCostos en columna Costo Total
    OLD_ROW = "${showUSD ? fU(d.usd_total) : '—'}</td>"
    NEW_ROW = "${showUSD ? fU((d.usd_total||0) + otrosCostos) : '—'}</td>"
    if OLD_ROW in html:
        html = html.replace(OLD_ROW, NEW_ROW, 1)
        fixes_ok.append("fila costos: Costo Total USD incluye otrosCostos")
    elif NEW_ROW in html:
        fixes_ok.append("fila costos: Costo Total USD (ya OK)")
    else:
        fixes_fail.append("fila costos: Costo Total USD (patron no encontrado)")

    # FIX EGRCATS: orden correcto (costos primero, luego gastos) y sin auto-sort por monto
    # FIX EGRCATS: orden correcto via regex
    EGRCATS_DESEADO = """const egrCats = [
    {k:'horas',       label:'Horas técnicas', color:'#1F5BA6',src:'fdata'},
    {k:'viat',        label:'Viáticos',       color:'#0369A1',src:'fdata'},
    {k:'costos_otros',label:'Otros costos',   color:'#0891B2',src:'extra'},
    {k:'nom',         label:'Nómina admin',   color:'#E8A020',src:'fdata'},
    {k:'comisiones',  label:'Comisiones',     color:'#F59E0B',src:'extra'},
    {k:'gastos_op',   label:'Gastos op.',     color:'#D97706',src:'extra'},
    {k:'imp',         label:'Impuestos',      color:'#94A3B8',src:'fdata'},
    {k:'amex',        label:'AMEX',           color:'#7C3AED',src:'fdata'},
    {k:'cf',          label:'Costos fijos',   color:'#06B6D4',src:'fdata'},
    {k:'seguros',     label:'Seguros',        color:'#65A30D',src:'extra'},
    {k:'ptu',         label:'PTU',            color:'#92400E',src:'extra'},
    {k:'prestamo',    label:'Préstamo',       color:'#DC2626',src:'extra'},
    {k:'estfi',       label:'Est. Fiscal',    color:'#EF4444',src:'extra'},
    {k:'activo_fijo', label:'Activo Fijo',    color:'#6366F1',src:'extra'},
  ];"""
    OLD_SORT = "}).filter(function(c){ return c.v>0; }).sort((a,b)=>b.v-a.v);"
    NEW_SORT  = "}).filter(function(c){ return c.v>0; });"
    import re as _re
    _egr_pattern = r'const egrCats = \[[\s\S]*?\];'
    if _re.search(_egr_pattern, html):
        html = _re.sub(_egr_pattern, EGRCATS_DESEADO, html, count=1)
        if OLD_SORT in html:
            html = html.replace(OLD_SORT, NEW_SORT, 1)
        fixes_ok.append("egrCats orden fijo (horas→nom→comisiones→gastos)")
    else:
        fixes_fail.append("egrCats (bloque no encontrado)")

    # FIX FCLIENTES_MES DUPLICADO: eliminar la segunda declaracion (formato con comillas simples, solo hasta AGO)
    # La segunda sobrescribe la primera (que tiene SEP) causando TOP 8 CLIENTES vacio en SEP
    _fcli_positions = [m.start() for m in _re.finditer(r'var FCLIENTES_MES\s*=\s*\{', html)]
    if len(_fcli_positions) >= 2:
        _idx2 = _fcli_positions[1]
        _chunk2 = html[_idx2:]
        _depth2 = 0; _ci = _chunk2.find('{')
        while _ci < len(_chunk2):
            if _chunk2[_ci] == '{': _depth2 += 1
            elif _chunk2[_ci] == '}':
                _depth2 -= 1
                if _depth2 == 0: break
            _ci += 1
        _end2 = _idx2 + _ci + 1
        if html[_end2:_end2+2] in (';\n', '; '):
            _end2 += 2
        elif html[_end2] == ';':
            _end2 += 1
        html = html[:_idx2] + html[_end2:]
        fixes_ok.append("FCLIENTES_MES duplicado eliminado (TOP8 SEP fix)")
    elif len(_fcli_positions) == 1:
        fixes_ok.append("FCLIENTES_MES OK (sin duplicado)")
    else:
        fixes_fail.append("FCLIENTES_MES (no encontrado)")

    # FIX CLIENTES MES: mostrar todos los clientes (sin limite), excluir egresos, incluir TRASPASO
    # Cambios: quitar slice(0,8), excluir solo egresos, colores con modulo, titulo correcto
    import re as _re2
    # a) Excluir egresos mezclados (sin TRASPASO que es ingreso real)
    _CLI_EXCLUIR_OLD = ("var _cliExcluir = ['GASTOS','SEGUROS','COSTOS FIJOS','IMPUESTOS','PRESTAMO','TRASPASO',"
                        "'EST FI','VENTA ACTIVO FIJO','N\u00d3MINA','NOMINA','COMISIONES','VI\u00c1TICOS','VIATICOS','PTU'];")
    _CLI_EXCLUIR_NEW = ("var _cliExcluir = ['GASTOS','SEGUROS','COSTOS FIJOS','IMPUESTOS','PRESTAMO',"
                        "'EST FI','VENTA ACTIVO FIJO','N\u00d3MINA','NOMINA','COMISIONES','VI\u00c1TICOS','VIATICOS','PTU'];")
    # b) Quitar .slice(0,8)
    _SLICE_OLD = "const cliEntries = Object.entries(cliFiltered).sort((a,b)=>b[1]-a[1]).slice(0,8);"
    _SLICE_NEW = "const cliEntries = Object.entries(cliFiltered).sort((a,b)=>b[1]-a[1]);"
    # c) Colores con modulo para mas de 18 clientes
    _COL_OLD = "colors8[i]+"
    _COL_NEW = "colors8[i%colors8.length]+"
    # d) Titulo
    _TTL_OLD = ">Top 8 Clientes \u2014 YTD</div>"
    _TTL_NEW = ">Clientes del mes</div>"
    # e) Label total
    _TOT_OLD = "Total Top 8:"
    _TOT_NEW = "Total clientes:"
    # f) Filtro base (si viene sin filtro aun)
    _CLI_BASE_OLD = ('var cliFiltered = {};\n'
                     '  months.forEach(function(m){\n'
                     '    if(FCLIENTES_MES[m]){\n'
                     '      Object.entries(FCLIENTES_MES[m]).forEach(function(e){\n'
                     '        const name=e[0], val=(typeof e[1]===\'object\' ? (e[1].mx||0) : (e[1]||0));\n'
                     '        cliFiltered[name] = (cliFiltered[name]||0) + val;\n'
                     '      });\n'
                     '    }\n'
                     '  });')
    _CLI_BASE_NEW = ('var cliFiltered = {};\n'
                     "  var _cliExcluir = ['GASTOS','SEGUROS','COSTOS FIJOS','IMPUESTOS','PRESTAMO',"
                     "'EST FI','VENTA ACTIVO FIJO','N\u00d3MINA','NOMINA','COMISIONES','VI\u00c1TICOS','VIATICOS','PTU'];\n"
                     '  months.forEach(function(m){\n'
                     '    if(FCLIENTES_MES[m]){\n'
                     '      Object.entries(FCLIENTES_MES[m]).forEach(function(e){\n'
                     '        const name=e[0], val=(typeof e[1]===\'object\' ? (e[1].mx||0) : (e[1]||0));\n'
                     '        if(_cliExcluir.indexOf(name) === -1) {\n'
                     '          cliFiltered[name] = (cliFiltered[name]||0) + val;\n'
                     '        }\n'
                     '      });\n'
                     '    }\n'
                     '  });')
    applied = []
    if _CLI_EXCLUIR_OLD in html:
        html = html.replace(_CLI_EXCLUIR_OLD, _CLI_EXCLUIR_NEW, 1); applied.append("excluir-sin-traspaso")
    if _CLI_BASE_OLD in html:
        html = html.replace(_CLI_BASE_OLD, _CLI_BASE_NEW, 1); applied.append("filtro-base")
    if _SLICE_OLD in html:
        html = html.replace(_SLICE_OLD, _SLICE_NEW, 1); applied.append("sin-slice8")
    html = html.replace(_COL_OLD, _COL_NEW)
    if _TTL_OLD in html:
        html = html.replace(_TTL_OLD, _TTL_NEW, 1); applied.append("titulo")
    if _TOT_OLD in html:
        html = html.replace(_TOT_OLD, _TOT_NEW); applied.append("total-label")
    if applied:
        fixes_ok.append("Clientes mes: " + "+".join(applied))
    else:
        fixes_ok.append("Clientes mes: ya OK")




    # FIX PASIVOS-CONST-2: cambiar segunda declaracion 'const PASIVOS' a 'var PASIVOS'
    # La regeneracion inserta 'const PASIVOS' dos veces (una por cada bloque duplicado).
    # En el mismo scope JS no se puede declarar const dos veces → SyntaxError L4xxx.
    # La solucion es cambiar la segunda instancia a 'var' que si permite redeclaracion.
    _pc_idx1 = html.find('const PASIVOS')
    if _pc_idx1 >= 0:
        _pc_idx2 = html.find('const PASIVOS', _pc_idx1 + 1)
        if _pc_idx2 >= 0:
            html = html[:_pc_idx2] + 'var PASIVOS' + html[_pc_idx2 + len('const PASIVOS'):]
            fixes_ok.append('PASIVOS-CONST-2 corregido (const→var)')
        else:
            fixes_ok.append('PASIVOS-CONST-2 (solo una instancia - OK)')
    else:
        fixes_ok.append('PASIVOS-CONST-2 (var PASIVOS — ya OK, no hay const)')

    # FIX HISTORIAL-PASIVOS-2: eliminar COMPLETAMENTE la segunda seccion de HISTORIAL PASIVOS
    # La segunda instancia (dentro del bloque del return ' sin cerrar en Chrome) tiene ~375 lineas de HTML
    # que son la causa raiz del SyntaxError L3295. El threshold de Chrome es ~702 lineas.
    # Sin esta seccion el bloque queda en ~478 lineas (bien por debajo del umbral).
    # La primera instancia (fuera del script, antes de L721) conserva sus datos y es la que se muestra.
    # IMPORTANTE: Se elimina desde <!-- HISTORIAL PASIVOS SECTION --> hasta <!-- END HISTORIAL PASIVOS -->
    # incluyendo el <script> de showRecursoMes que esta dentro de esa seccion.
    _HP_MARKER = '<!-- HISTORIAL PASIVOS SECTION -->'
    _HP_END = '<!-- END HISTORIAL PASIVOS -->'
    _hp1 = html.find(_HP_MARKER)
    _hp2 = -1
    if _hp1 >= 0:
        _hp2 = html.find(_HP_MARKER, _hp1 + 1)
    if _hp2 >= 0:
        # Hay dos instancias - eliminar la segunda completa (incluyendo END marker)
        _hp_end_idx = html.find(_HP_END, _hp2)
        if _hp_end_idx >= 0:
            _hp_end_full = _hp_end_idx + len(_HP_END)
            html = html[:_hp2] + html[_hp_end_full:]
            fixes_ok.append('HISTORIAL-PASIVOS-2 eliminado')
        else:
            fixes_fail.append('HISTORIAL-PASIVOS-2 (END marker no encontrado)')
    elif _hp1 >= 0:
        fixes_ok.append('HISTORIAL-PASIVOS-2 (segunda instancia ya no existe - OK)')
    else:
        fixes_fail.append('HISTORIAL-PASIVOS-2 (marker HISTORIAL PASIVOS SECTION no encontrado)')

    # FIX VIEW-FLUJO-NESTING: la seccion "HISTORIAL DE PASIVOS" (div.section abierto en L620)
    # nunca se cierra antes del cierre de view-costos. Esto hace que view-flujo quede anidado
    # DENTRO de view-costos, y al poner view-costos en display:none el flujo desaparece.
    # Solucion: agregar </div><!-- /section historial pasivos --> antes de </div><!-- /view-costos -->
    _VF_OLD = '<!-- END HISTORIAL PASIVOS -->\n</div><!-- /view-costos -->'
    _VF_NEW = '<!-- END HISTORIAL PASIVOS -->\n</div><!-- /section historial pasivos -->\n</div><!-- /view-costos -->'
    if _VF_OLD in html:
        html = html.replace(_VF_OLD, _VF_NEW)
        fixes_ok.append('VIEW-FLUJO-NESTING corregido (section div cerrado antes de view-costos)')
    elif _VF_NEW in html:
        fixes_ok.append('VIEW-FLUJO-NESTING (ya OK)')
    else:
        fixes_fail.append('VIEW-FLUJO-NESTING (patron END HISTORIAL PASIVOS + /view-costos no encontrado)')

    # FIX COSTOS-TABLE: si el bloque de filtro de costos está truncado (sin ctbody HTML),
    # restaurar la tabla completa. El ctbody puede existir en JS pero no como elemento HTML.
    import re as _re_ct
    _CT_HTML_CHECK = '<tbody[^>]+id=.ctbody.'
    _ct_html_ok = bool(_re_ct.search(_CT_HTML_CHECK, html))
    if not _ct_html_ok:
        # Anchor A: comentario AGRUPAR POR vacío (estado actual tras regeneración)
        _CT_ANCHOR_A = '  <!-- AGRUPAR POR -->\n  <!-- /section costos desglose -->'
        # Anchor B: filter-wrap truncado (estado anterior)
        _CT_ANCHOR_B = '  <div class="filter-wrap" style="margin-bottom:20px;justify-content:space-between">\n  <!-- /section-alt -->'
        _CT_TABLE_BLOCK = """  <!-- AGRUPAR POR -->
  <div style="display:flex;gap:8px;margin-bottom:12px;align-items:center">
    <span style="font-size:.65rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#6B7A99">AGRUPAR POR:</span>
    <button onclick="setCostosGroup('cliente',this)" id="cgrp-cliente" style="padding:4px 14px;border-radius:20px;border:none;background:#0D1F3C;color:#fff;font-size:.72rem;cursor:pointer;font-weight:600">Cliente</button>
    <button onclick="setCostosGroup('po',this)" id="cgrp-po" style="padding:4px 14px;border-radius:20px;border:1px solid #D1D5DB;background:#F9FAFB;color:#374151;font-size:.72rem;cursor:pointer">PO</button>
    <button onclick="setCostosGroup('servicio',this)" id="cgrp-servicio" style="padding:4px 14px;border-radius:20px;border:1px solid #D1D5DB;background:#F9FAFB;color:#374151;font-size:.72rem;cursor:pointer">Servicio</button>
  </div>
  <div style="margin-bottom:14px">
    <input id="costos-search" type="text" placeholder="Buscar cliente, PO o servicio..." oninput="buildCostosTable()" style="width:100%;padding:8px 14px;border:1px solid #E5E7EB;border-radius:8px;font-size:.78rem;outline:none;box-sizing:border-box">
  </div>
  <div style="overflow-x:auto;border-radius:10px;border:1px solid #E5E7EB">
    <table style="width:100%;border-collapse:collapse;font-size:.78rem">
      <thead>
        <tr style="background:#0D1F3C!important">
          <th style="padding:10px 14px;text-align:left;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;min-width:120px;border-bottom:none">Mes</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none">Costo Hrs USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none">Perdiem USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none">Otros Costos USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none">Costo Total USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none">Costo MXN</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none"></th>
        </tr>
      </thead>
      <tbody id="ctbody"></tbody>
    </table>
  </div>
  <!-- /section costos desglose -->"""
        if _CT_ANCHOR_A in html:
            html = html.replace(_CT_ANCHOR_A, _CT_TABLE_BLOCK, 1)
            fixes_ok.append('COSTOS-TABLE restaurada con ctbody (anchor A)')
        elif _CT_ANCHOR_B in html:
            html = html.replace(_CT_ANCHOR_B, _CT_TABLE_BLOCK + '\n<!-- /section-alt -->', 1)
            fixes_ok.append('COSTOS-TABLE restaurada con ctbody (anchor B)')
        else:
            fixes_fail.append('COSTOS-TABLE (patron no encontrado — revisar anchor)')
    else:
        fixes_ok.append('COSTOS-TABLE (ctbody ya OK)')

    # FIX CANVAS-COSTOS: si faltan los canvas de comportamiento mensual de costos, insertarlos
    _CANVAS_CHECK = 'costos-tendencia-chart'
    if _CANVAS_CHECK not in html:
        _CANVAS_ANCHOR = '  <!-- AGRUPAR POR -->'
        _CANVAS_INSERT = """  <!-- GRAFICAS COSTOS -->
  <div class="chart-grid-2" style="margin-bottom:20px">
    <div class="chart-card">
      <div class="chart-title"><span class="dot" style="background:#1F5BA6"></span>Tendencia Mensual vs Acumulado</div>
      <canvas id="costos-tendencia-chart" style="max-height:220px"></canvas>
    </div>
    <div class="chart-card">
      <div class="chart-title"><span class="dot" style="background:#F59E0B"></span>Costo Total USD por Mes</div>
      <canvas id="costos-barras-chart" style="max-height:220px"></canvas>
    </div>
  </div>

  <!-- AGRUPAR POR -->"""
        idx_ca = html.find(_CANVAS_ANCHOR)
        if idx_ca >= 0:
            html = html.replace(_CANVAS_ANCHOR, _CANVAS_INSERT, 1)
            fixes_ok.append('CANVAS-COSTOS insertados')
        else:
            fixes_fail.append('CANVAS-COSTOS (ancla AGRUPAR POR no encontrada)')
    else:
        fixes_ok.append('CANVAS-COSTOS (ya OK)')

    # FIX SALDO-DIA: siempre leer el último saldo real de bancos del Excel
    # y escribirlo en var SALDO_DIA (con fecha del último dato y saldos por cuenta)
    import json as _json
    _saldo_data = None
    try:
        _saldo_data = extraer_saldo_bancos()
    except Exception as _e:
        print(f"  WARN extraer_saldo_bancos: {_e}")

    if _saldo_data is None:
        _saldo_data = {'fecha': None, 'total': 0, 'banks': {}}

    _SD_JS = 'var SALDO_DIA = ' + _json.dumps(_saldo_data, ensure_ascii=False) + ';'
    _SD_ANCHOR = 'function fRenderSaldoDia(){'

    if 'var SALDO_DIA' in html:
        # Reemplazar el existente (puede ser viejo o el fallback del turno anterior)
        import re as _re
        html = _re.sub(r'var SALDO_DIA\s*=\s*\{[^;]+\};', _SD_JS, html, count=1)
        fixes_ok.append('SALDO-DIA actualizado con datos reales')
    else:
        idx_sd = html.find(_SD_ANCHOR)
        if idx_sd >= 0:
            html = html[:idx_sd] + _SD_JS + '\n' + html[idx_sd:]
            fixes_ok.append('SALDO-DIA insertado con datos reales')
        else:
            fixes_fail.append('SALDO-DIA (function fRenderSaldoDia no encontrado)')

    # FIX: fRenderSaldoDia mostrar "Actualizado: <hoy> · Datos al: <fecha_excel>"
    _FECHA_OLD = "    elFecha.textContent = 'Saldo al '+hoyStr + (s.fecha ? ' · último dato: '+s.fecha : '');"
    _FECHA_NEW = "    elFecha.textContent = 'Actualizado: '+(s.fecha_generacion||hoyStr) + (s.fecha ? ' · Datos al: '+s.fecha : '');"
    if _FECHA_OLD in html:
        html = html.replace(_FECHA_OLD, _FECHA_NEW, 1)
        fixes_ok.append('FECHA-RENDER actualizado')
    elif _FECHA_NEW in html:
        fixes_ok.append('FECHA-RENDER (ya OK)')
    else:
        fixes_fail.append('FECHA-RENDER (patron no encontrado)')

    # FIX: thead de costos — 7 columnas con background #0D1F3C!important en cada <th>
    # (el CSS global "th{background:var(--gray-100)}" sobreescribia el color del <tr>)
    _CTHEAD_OLD = """      <thead>
        <tr style="background:#0D1F3C">
          <th style="padding:10px 14px;text-align:left;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff;min-width:120px">Mes</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff">Costo Hrs USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff">Perdiem USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff">Otros Costos USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff">Costo Total</th>
        </tr>
      </thead>
      <tbody id="ctbody"></tbody>"""
    _CTHEAD_OLD2 = """      <thead>
        <tr style="background:#0D1F3C">
          <th style="padding:10px 14px;text-align:left;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff;min-width:120px">Mes</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff">Costo Hrs USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff">Perdiem USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff">Otros Costos USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff">Costo Total USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff">Costo MXN</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff"></th>
        </tr>
      </thead>
      <tbody id="ctbody"></tbody>"""
    _CTHEAD_NEW = """      <thead>
        <tr style="background:#0D1F3C!important">
          <th style="padding:10px 14px;text-align:left;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;min-width:120px;border-bottom:none">Mes</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none">Costo Hrs USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none">Perdiem USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none">Otros Costos USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none">Costo Total USD</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none">Costo MXN</th>
          <th style="padding:10px 14px;text-align:right;font-size:.65rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#fff!important;background:#0D1F3C!important;border-bottom:none"></th>
        </tr>
      </thead>
      <tbody id="ctbody"></tbody>"""
    if _CTHEAD_OLD in html:
        html = html.replace(_CTHEAD_OLD, _CTHEAD_NEW, 1)
        fixes_ok.append('COSTOS-THEAD azul navy con !important')
    elif _CTHEAD_OLD2 in html:
        html = html.replace(_CTHEAD_OLD2, _CTHEAD_NEW, 1)
        fixes_ok.append('COSTOS-THEAD azul navy con !important')
    elif 'background:#0D1F3C!important' in html and 'id="ctbody"' in html:
        fixes_ok.append('COSTOS-THEAD (ya OK con !important)')
    else:
        fixes_fail.append('COSTOS-THEAD (patron no encontrado)')

    # FIX: bankDefs — 5 bancos desde CONCENTRADO, con isUSD para mostrar sub-monto en dólares
    # banks tiene objetos {mxn, usd?} en vez de números simples
    _BANKDEFS_MARKER_OK = "key:'BBVA USD',   label:'BBVA USD',   isUSD:true"
    _BANKDEFS_MARKER_OK2 = "key:'BBVA USD',   label:'BBVA USD',   isUSD:true"
    _BANKDEFS_OLD_4_ISMXN = """  var bankDefs = [
    {key:'BBVA Pesos', label:'BBVA Pesos', isMXN:true},
    {key:'BBVA USD',   label:'BBVA USD',   isMXN:false},
    {key:'BMX USD',    label:'BMX USD',    isMXN:false},
    {key:'BMX MXN',    label:'BMX MXN',    isMXN:true}
  ];"""
    _BANKDEFS_OLD_5_VSBVA = """  var bankDefs = [
    {key:'Vs BBVA pesos', label:'BBVA Pesos', isMXN:true},
    {key:'Vs BBVA USD',   label:'BBVA USD',   isMXN:false},
    {key:'BMX USD',       label:'BMX USD',    isMXN:false},
    {key:'BMXQRO MN',     label:'BMX QRO MN', isMXN:true},
    {key:'BANAMEX MN',    label:'Banamex MN', isMXN:true}
  ];"""
    _BANKDEFS_NEW_5 = """  var bankDefs = [
    {key:'BBVA Pesos', label:'BBVA Pesos', isUSD:false},
    {key:'BBVA USD',   label:'BBVA USD',   isUSD:true},
    {key:'BMX USD',    label:'BMX USD',    isUSD:true},
    {key:'BMX QRO MN', label:'BMX QRO MN', isUSD:false},
    {key:'Banamex MN', label:'Banamex MN', isUSD:false}
  ];"""
    if _BANKDEFS_MARKER_OK in html:
        fixes_ok.append('BANKDEFS (ya OK - 5 bancos con isUSD)')
    elif _BANKDEFS_OLD_4_ISMXN in html:
        html = html.replace(_BANKDEFS_OLD_4_ISMXN, _BANKDEFS_NEW_5, 1)
        fixes_ok.append('BANKDEFS actualizado (4->5 bancos con isUSD)')
    elif _BANKDEFS_OLD_5_VSBVA in html:
        html = html.replace(_BANKDEFS_OLD_5_VSBVA, _BANKDEFS_NEW_5, 1)
        fixes_ok.append('BANKDEFS actualizado (5 bancos Vs->/display con isUSD)')
    else:
        fixes_fail.append('BANKDEFS (patron no encontrado)')

    # FIX: Renderizado de tarjetas bancarias — MXN principal + USD abajo para cuentas USD
    _CARDS_OLD = """  var cards = bankDefs.map(function(b){
    var v = s.banks[b.key] || 0;
    var vFmt;
    if(v>=1e6) vFmt='$'+(v/1e6).toFixed(1)+'M';
    else if(v>=1e3) vFmt='$'+(v/1e3).toFixed(0)+'K';
    else vFmt='$'+Math.round(v);
    var sub = b.isMXN ? 'MXN' : 'MXN equiv.';
    return '<div style="background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.08);border-radius:7px;padding:7px 9px">'
      +'<div style="font-size:.52rem;font-weight:700;color:#93C5FD;text-transform:uppercase;letter-spacing:.06em;margin-bottom:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+b.label+'</div>'
      +'<div style="font-size:.72rem;font-weight:800;color:#fff;white-space:nowrap">'+vFmt+'</div>'
      +'<div style="font-size:.5rem;color:#64748B;margin-top:1px">'+sub+'</div>'
      +'</div>';
  });
  elBanks.innerHTML = cards.join('');"""
    _CARDS_NEW = """  var fmtMXN=function(v){if(v>=1e6)return '$'+(v/1e6).toFixed(2).replace(/\.?0+$/,'')+'M';if(v>=1e3)return '$'+(v/1e3).toFixed(0)+'K';return '$'+Math.round(v);};
  var fmtUSD=function(v){if(v>=1e6)return 'USD '+(v/1e6).toFixed(2).replace(/\.?0+$/,'')+'M';if(v>=1e3)return 'USD '+(v/1e3).toFixed(0)+'K';return 'USD '+Math.round(v);};
  var cards = bankDefs.map(function(b){
    var bv=s.banks[b.key]||{};
    var mxn=(typeof bv==='object'&&bv!==null&&'mxn' in bv)?bv.mxn:(typeof bv==='number'?bv:0);
    var usd=(typeof bv==='object'&&bv!==null&&'usd' in bv)?bv.usd:null;
    var usdLine=(b.isUSD&&usd!==null)?'<div style="font-size:.5rem;color:#93C5FD;margin-top:2px;font-weight:600">'+fmtUSD(usd)+'</div>':'<div style="font-size:.5rem;color:#64748B;margin-top:1px">MXN</div>';
    return '<div style="background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.08);border-radius:7px;padding:7px 9px">'
      +'<div style="font-size:.52rem;font-weight:700;color:#93C5FD;text-transform:uppercase;letter-spacing:.06em;margin-bottom:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+b.label+'</div>'
      +'<div style="font-size:.72rem;font-weight:800;color:#fff;white-space:nowrap">'+fmtMXN(mxn)+' MXN</div>'
      +usdLine
      +'</div>';
  });
  elBanks.innerHTML = cards.join('');"""
    if 'fmtMXN=function' in html and 'fmtUSD=function' in html:
        fixes_ok.append('CARDS-RENDER (ya OK - MXN+USD)')
    elif _CARDS_OLD in html:
        html = html.replace(_CARDS_OLD, _CARDS_NEW, 1)
        fixes_ok.append('CARDS-RENDER actualizado (MXN principal + USD abajo)')
    else:
        fixes_fail.append('CARDS-RENDER (patron no encontrado)')

    # FIX: Grid de bancos — 5 columnas
    _GRID_OLD_4 = 'id="flujo-saldo-banks" style="display:grid;grid-template-columns:repeat(4,1fr);gap:6px"'
    _GRID_5 = 'id="flujo-saldo-banks" style="display:grid;grid-template-columns:repeat(5,1fr);gap:6px"'
    if _GRID_5 in html:
        fixes_ok.append('BANKS-GRID (ya OK - 5 cols)')
    elif _GRID_OLD_4 in html:
        html = html.replace(_GRID_OLD_4, _GRID_5, 1)
        fixes_ok.append('BANKS-GRID actualizado (4->5 cols)')

    # FIX COSTOS-SECTION-DIV: falta </div> de cierre del <div class="section"> de TABLA MENSUAL.
    # Sin este cierre, view-flujo queda anidado dentro de view-costos y aparece en blanco.
    # COSTOS-SECTION-DIV: verificar que view-flujo NO esta anidado dentro de view-costos
    # El indicador es que id="view-flujo" aparece DESPUES del cierre de id="view-costos"
    _vc_end = html.find('</div>', html.find('id="view-costos"'))
    _vf_pos = html.find('id="view-flujo"')
    if _vf_pos > 0 and _vc_end > 0 and _vf_pos > _vc_end:
        fixes_ok.append('COSTOS-SECTION-DIV (view-flujo correctamente despues de view-costos)')
    else:
        # Intentar fix con patrones conocidos
        _SEC_CLOSE_CHECK = '</div>\n<!-- /section-alt -->\n<!-- HISTORIAL PASIVOS SECTION -->'
        _SEC_CLOSE_MISSING = '<!-- /section costos desglose -->\n<!-- /section-alt -->\n<!-- HISTORIAL PASIVOS SECTION -->'
        _SEC_CLOSE_FIXED   = '<!-- /section costos desglose -->\n</div>\n<!-- /section-alt -->\n<!-- HISTORIAL PASIVOS SECTION -->'
        if _SEC_CLOSE_CHECK in html:
            fixes_ok.append('COSTOS-SECTION-DIV (cierre ya OK - patron clasico)')
        elif _SEC_CLOSE_MISSING in html:
            html = html.replace(_SEC_CLOSE_MISSING, _SEC_CLOSE_FIXED, 1)
            fixes_ok.append('COSTOS-SECTION-DIV: </div> restaurado')
        else:
            fixes_ok.append('COSTOS-SECTION-DIV (estructura no reconocida — verificar visualmente)')

    # -------------------------------------------------------------------------
    # FIX CRES VIATICOS: usar e.perdiem en lugar de diferencia usd_cost - usdHrsCost
    # Estos fixes aseguran que la columna VIÁTICOS muestre el costo real en la expansion
    # -------------------------------------------------------------------------

    # 1) Calculo de viat en expansion por worker (toggleCostosDetail) — ambos patrones (viejo y ya-ok)
    OLD_VIAT_CALC = "    const usdHrsCost = e.rate>0 ? e.rate * e.usd_hrs : 0;\n    const viat = e.usd_cost - usdHrsCost;"
    NEW_VIAT_CALC = "    const usdHrsCost = e.rate>0 ? e.rate * e.usd_hrs : 0;\n    const viat = (e.perdiem != null && e.perdiem > 0) ? e.perdiem : Math.max(0, e.usd_cost - usdHrsCost);"
    if OLD_VIAT_CALC in html:
        html = html.replace(OLD_VIAT_CALC, NEW_VIAT_CALC)
        fixes_ok.append('CRES viat: usar e.perdiem en expansion')
    elif NEW_VIAT_CALC in html:
        fixes_ok.append('CRES viat: expansion (ya OK)')
    else:
        fixes_fail.append('CRES viat: expansion (patron no encontrado)')

    # 2) Render columna VIATICOS en fila worker:
    #    a) Quitar showUSD para que siempre muestre perdiem (viáticos son USD siempre)
    #    b) COSTO TOTAL USD debe incluir perdiem: e.usd_cost + (e.perdiem||0)
    #    c) Sort incluye perdiem en la clave de orden
    # Este fix cubre el patron exacto del HTML actual (ya con e.perdiem en el calculo de viat)
    import re as _vre
    # Fix render viáticos: quitar showUSD && del render de viat
    _OLD_VIAT_SHOW = "${showUSD && viat>0.01 ? fU(viat) : '—'}</td>"
    _NEW_VIAT_SHOW = "${viat>0.01 ? fU(viat) : '—'}</td>"
    if _OLD_VIAT_SHOW in html:
        html = html.replace(_OLD_VIAT_SHOW, _NEW_VIAT_SHOW, 1)
        fixes_ok.append('CRES viat: quitar showUSD del render')
    elif _NEW_VIAT_SHOW in html:
        fixes_ok.append('CRES viat: render (ya sin showUSD)')
    else:
        fixes_fail.append('CRES viat: render showUSD (patron no encontrado)')

    # Fix COSTO TOTAL USD worker: incluir perdiem en el total del worker
    _OLD_TOTAL_W = "${showUSD && e.usd_cost>0 ? fU(e.usd_cost) : '—'}</td>\n      <td class=\"num\">${costMXN}</td>"
    _NEW_TOTAL_W = "${showUSD && (e.usd_cost+(e.perdiem||0))>0 ? fU(e.usd_cost+(e.perdiem||0)) : '—'}</td>\n      <td class=\"num\">${costMXN}</td>"
    if _OLD_TOTAL_W in html:
        html = html.replace(_OLD_TOTAL_W, _NEW_TOTAL_W, 1)
        fixes_ok.append('CRES total worker: incluye perdiem')
    elif _NEW_TOTAL_W in html:
        fixes_ok.append('CRES total worker: (ya OK)')
    else:
        fixes_fail.append('CRES total worker: (patron no encontrado)')

    # 3) Sort expansion: incluir perdiem en la clave de orden para workers solo-viat
    OLD_SORT_CRES = "entries.sort((a,b)=>(b.usd_cost+b.mxn_cost)-(a.usd_cost+a.mxn_cost))"
    NEW_SORT_CRES = "entries.sort((a,b)=>((b.usd_cost+(b.perdiem||0))+b.mxn_cost)-((a.usd_cost+(a.perdiem||0))+a.mxn_cost))"
    if OLD_SORT_CRES in html:
        html = html.replace(OLD_SORT_CRES, NEW_SORT_CRES)
        fixes_ok.append('CRES sort: incluye perdiem en orden')
    elif NEW_SORT_CRES in html:
        fixes_ok.append('CRES sort: (ya OK)')
    else:
        fixes_fail.append('CRES sort: (patron no encontrado)')

    # 4) Calculo de viat en export XLS por mes
    OLD_VIAT_XLS1 = "    const usdHrs = e.rate>0 ? e.rate*e.usd_hrs : 0;\n    const viat = e.usd_cost - usdHrs;\n    rows.push([e.n"
    NEW_VIAT_XLS1 = "    const usdHrs = e.rate>0 ? e.rate*e.usd_hrs : 0;\n    const viat = (e.perdiem != null && e.perdiem > 0) ? e.perdiem : Math.max(0, e.usd_cost - usdHrs);\n    rows.push([e.n"
    if OLD_VIAT_XLS1 in html:
        html = html.replace(OLD_VIAT_XLS1, NEW_VIAT_XLS1)
        fixes_ok.append('CRES viat: export XLS mes')
    elif NEW_VIAT_XLS1 in html:
        fixes_ok.append('CRES viat: export XLS mes (ya OK)')

    # 5) Calculo de viat en export XLS todos los meses
    OLD_VIAT_XLS2 = "        const usdHrs = e.rate>0 ? e.rate*(e.usd_hrs||0) : 0;\n        const viat = e.usd_cost - usdHrs;\n        rows.push([mes"
    NEW_VIAT_XLS2 = "        const usdHrs = e.rate>0 ? e.rate*(e.usd_hrs||0) : 0;\n        const viat = (e.perdiem != null && e.perdiem > 0) ? e.perdiem : Math.max(0, e.usd_cost - usdHrs);\n        rows.push([mes"
    if OLD_VIAT_XLS2 in html:
        html = html.replace(OLD_VIAT_XLS2, NEW_VIAT_XLS2)
        fixes_ok.append('CRES viat: export XLS todos')
    elif NEW_VIAT_XLS2 in html:
        fixes_ok.append('CRES viat: export XLS todos (ya OK)')

    # 6) Badge $x/día: usar e.rate_viat del CRES (valor del mes) en lugar del mapa global RATE_VIAT
    OLD_BADGE_VIAT_V1 = "const rateViat = (RATE_VIAT && RATE_VIAT[e.n.toUpperCase()]) || 0;"
    OLD_BADGE_VIAT = "const rateViat = (e.rate_viat && e.rate_viat > 0) ? e.rate_viat : ((RATE_VIAT && RATE_VIAT[e.n.toUpperCase()]) || 0);"
    NEW_BADGE_VIAT = "const rateViat = (e.rate_viat != null && e.rate_viat > 0) ? e.rate_viat : 0;"
    cnt_badge = html.count(OLD_BADGE_VIAT_V1) + html.count(OLD_BADGE_VIAT)
    if html.count(OLD_BADGE_VIAT_V1) > 0:
        html = html.replace(OLD_BADGE_VIAT_V1, NEW_BADGE_VIAT)
        fixes_ok.append(f"CRES badge $/día: usa e.rate_viat del mes (v1, {html.count(NEW_BADGE_VIAT)} ocurrencias)")
    elif html.count(OLD_BADGE_VIAT) > 0:
        html = html.replace(OLD_BADGE_VIAT, NEW_BADGE_VIAT)
        fixes_ok.append(f"CRES badge $/día: usa e.rate_viat del mes (v2, {html.count(NEW_BADGE_VIAT)} ocurrencias)")
    elif NEW_BADGE_VIAT in html:
        fixes_ok.append('CRES badge $/día: (ya OK, usa e.rate_viat)')
    else:
        fixes_fail.append('CRES badge $/día: patron no encontrado')

    # 9) cFilterMes: limpiar search y actualizar estilos inline
    _CFILTER_MARKER = '// Limpiar search al cambiar mes'
    if _CFILTER_MARKER not in html:
        _OLD_CF = 'function cFilterMes(m){'
        _NEW_CF = ('function cFilterMes(m){\n'
            '  // Limpiar search al cambiar mes\n'
            "  const _cs = document.getElementById('costos-search'); if(_cs) _cs.value='';\n"
            '  // Actualizar estilos inline de botones para que se vean bien\n'
            "  document.querySelectorAll('.cfilter-btn').forEach(b=>{\n"
            '    const isActive = b.dataset.m === m;\n'
            "    b.style.background = isActive ? 'var(--blue)' : '';\n"
            "    b.style.color = isActive ? '#fff' : '';\n"
            "    b.style.border = isActive ? 'none' : '';\n"
            '  });')
        if _OLD_CF in html:
            html = html.replace(_OLD_CF, _NEW_CF, 1)
            fixes_ok.append('cFilterMes: estilos inline y limpia search')
        else:
            fixes_fail.append('cFilterMes: patron no encontrado')
    else:
        fixes_ok.append('cFilterMes: (ya OK)')

    print("  FIXES OK:", ", ".join(fixes_ok))
    if fixes_fail:
        print("  FIXES FALLIDOS:", ", ".join(fixes_fail))
    return html

def actualizar_pasivo_html(html, cierres, detalle, vivo):
    """Actualiza la sección hardcodeada de HISTORIAL PASIVOS en el HTML."""
    print("  Actualizando HISTORIAL PASIVOS en HTML...")

    fN = lambda n: f"{n:,.1f}" if n != int(n) else f"{int(n):,}"
    fU = lambda n: f"${int(round(n)):,}"

    # 1. Actualizar badge "Cierre vivo"
    if vivo:
        import re
        html = re.sub(
            r'⚡ Cierre vivo: [a-z]{3}-\d{2,4}',
            f'⚡ Cierre vivo: {vivo}',
            html
        )

    # 2. Actualizar KPI cards del cierre vivo
    if vivo:
        m_vivo = {'ene-26':'ENE','feb-26':'FEB','mar-26':'MAR','abr-26':'ABR',
                  'may-26':'MAY','jun-26':'JUN','jul-26':'JUL','ago-26':'AGO','sep-26':'SEP'}.get(vivo)
        if m_vivo and m_vivo in cierres:
            d = cierres[m_vivo]
            n_clientes = len(detalle.get(m_vivo, []))
            import re
            # Personas
            html = re.sub(r'(<div class="kpi-value" style="color:var\(--navy\)">)\d+(</div>\s*<div class="kpi-sub">Recursos)',
                          lambda x: f'{x.group(1)}{int(d["personas"])}{x.group(2)}', html)
            # Horas pendientes
            html = re.sub(r'(<div class="kpi-value" style="color:var\(--blue\)">)[0-9,.]+(</div>)',
                          lambda x: f'{x.group(1)}{fN(d["horas"])}{x.group(2)}', html, count=1)
            # Total USD
            html = re.sub(r'(<div class="kpi-value" style="color:var\(--green\)">\$)[0-9,]+(</div>\s*<div class="kpi-sub">TC:)',
                          lambda x: f'{x.group(1)}{int(round(d["total_usd"])):,}{x.group(2)}', html)
            # TC en sub
            html = re.sub(r'(TC: )[0-9.]+', f'TC: {d["tc"]:.4f}', html, count=1)
            # Neto MXN
            html = re.sub(r'(<div class="kpi-value" style="color:var\(--gold\)">\$)[0-9,]+(</div>)',
                          lambda x: f'{x.group(1)}{int(round(d["neto_mxn"])):,}{x.group(2)}', html, count=1)
            # Neto USD
            html = re.sub(r'(<div class="kpi-value" style="color:var\(--green\)">\$)[0-9,]+(</div>\s*<div class="kpi-sub">MATUK)',
                          lambda x: f'{x.group(1)}{int(round(d["neto_usd"])):,}{x.group(2)}', html)
            # Clientes
            html = re.sub(r'(Con horas pendientes )[a-z]{3}-\d{2,4}', f'Con horas pendientes {vivo}', html)

    # 3. Reconstruir tabla CIERRES HISTÓRICOS
    orden_meses = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP']
    MES_LABEL = {'ENE':'ene-26','FEB':'feb-26','MAR':'mar-26','ABR':'abr-26',
                 'MAY':'may-26','JUN':'jun-26','JUL':'jul-26','AGO':'ago-26','SEP':'sep-26'}
    m_vivo_code = {'ene-26':'ENE','feb-26':'FEB','mar-26':'MAR','abr-26':'ABR',
                   'may-26':'MAY','jun-26':'JUN','jul-26':'JUL','ago-26':'AGO','sep-26':'SEP'}.get(vivo,'')

    rows_html = ''
    alt = False
    for m in orden_meses:
        if m not in cierres: continue
        if m == m_vivo_code: continue  # cierre vivo no va en históricos
        d = cierres[m]
        bg = '#F9FAFB' if alt else '#fff'
        alt = not alt
        neto_usd_str = fU(d['neto_usd']) if d['neto_usd'] > 0 else '—'
        neto_usd_color = '#059669' if d['neto_usd'] > 0 else '#9CA3AF'
        rows_html += f'''          <tr style="background:{bg};border-bottom:1px solid #F3F4F6">
            <td style="padding:10px 14px;font-weight:600;color:var(--navy)">{MES_LABEL[m]}</td>
            <td style="padding:10px 14px;text-align:right;font-variant-numeric:tabular-nums">{d["tc"]:.4f}</td>
            <td style="padding:10px 14px;text-align:right;font-variant-numeric:tabular-nums">{int(d["personas"])}</td>
            <td style="padding:10px 14px;text-align:right;font-variant-numeric:tabular-nums;color:#1F5BA6;font-weight:600">{fN(d["horas"])}</td>
            <td style="padding:10px 14px;text-align:right;font-variant-numeric:tabular-nums;color:#059669;font-weight:600">{fU(d["total_usd"])}</td>
            <td style="padding:10px 14px;text-align:right;font-variant-numeric:tabular-nums">{fU(d["total_mxn"])}</td>
            <td style="padding:10px 14px;text-align:right;font-variant-numeric:tabular-nums">{fU(d["neto_mxn"])}</td>
            <td style="padding:10px 14px;text-align:right;font-variant-numeric:tabular-nums;color:{neto_usd_color}">{neto_usd_str}</td>
          </tr>\n'''

    # Reemplazar tbody de cierres históricos (funciona con tbody vacío o con filas)
    import re
    # Intentar con filas existentes primero
    new_html, n = re.subn(
        r'(<tbody>\s*)(?:<tr[\s\S]*?</tr>\s*)+(</tbody>)',
        lambda x: x.group(1) + rows_html + x.group(2),
        html, count=1
    )
    if n:
        html = new_html
    else:
        # tbody completamente vacío
        for pat in ['<tbody>\n          </tbody>', '<tbody>\n        </tbody>', '<tbody></tbody>']:
            if pat in html:
                html = html.replace(pat, '<tbody>\n' + rows_html + pat[7:], 1)
                break

    # 4. Reconstruir PASIVOS JS y botones de tabs
    tabs_html = ''
    pasivos_js = {}
    meses_disponibles = [m for m in orden_meses if m in detalle and detalle[m]]
    ultimo_mes = meses_disponibles[-1] if meses_disponibles else None
    for m in meses_disponibles:
        label = MES_LABEL[m]
        btn_id = f"pbtn-{label[:3]}"
        # Botón activo = el ÚLTIMO mes (más reciente)
        if m == ultimo_mes:
            style = 'background:var(--blue);color:#fff;border:1px solid var(--blue)'
        else:
            style = 'border:1px solid #CBD5E1;background:#F1F5F9;color:#64748B'
        tabs_html += f'      <button onclick="showPasivoMes(\'{label}\',this)" id="{btn_id}" style="padding:5px 12px;border-radius:6px;{style};font-size:.75rem;cursor:pointer">{label.upper()}</button>\n'
        pasivos_js[label] = detalle[m]

    # Reemplazar tabs (funciona con div vacío o con botones existentes)
    import re as _re
    new_html, n = _re.subn(
        r'(<div style="display:flex;gap:8px;margin-bottom:8px">)\s*(?:<button[\s\S]*?</button>\s*)*(</div>)',
        lambda x: x.group(1) + '\n' + tabs_html + '    ' + x.group(2),
        html, count=1
    )
    if n:
        html = new_html
    else:
        print('    ⚠️  No se encontró div de tabs para reemplazar')

    # Reemplazar const PASIVOS
    pasivos_str = json.dumps(pasivos_js, ensure_ascii=False)
    html = re.sub(
        r'const PASIVOS = \{[\s\S]*?\};',
        f'const PASIVOS = {pasivos_str};',
        html, count=1
    )

    # Actualizar showPasivoMes call en showTab con el primer mes disponible
    first_mes = MES_LABEL.get([m for m in orden_meses if m in detalle and detalle[m]][0] if any(detalle.get(m) for m in orden_meses) else 'JUL', 'jul-26')
    first_btn_id = f"pbtn-{first_mes[:3]}"
    html = re.sub(
        r"showPasivoMes\('[a-z]{3}-\d{2}', document\.getElementById\('pbtn-[a-z]{3}'\)\)",
        f"showPasivoMes('{first_mes}', document.getElementById('{first_btn_id}'))",
        html, count=1
    )

    print(f"    OK HISTORIAL PASIVOS actualizado — meses: {[MES_LABEL[m] for m in orden_meses if m in cierres]}")
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
    horas_data, breakdown, resources, cres = extraer_horas()
    costos                  = extraer_costos(tc_mes)
    cierres, detalle, vivo  = extraer_pasivo()
    fegr_extra              = extraer_fegr_extra()

    print()
    for m in MESES:
        d=horas_data[m]
        c=costos.get(m,{})
        print(f"  {m}: {round(d['hours'])}h  ING_MXN={total_mes[m]:>12,.0f}  COSTOS_USD={c.get('usd_total',0):>9,.0f}")
    print()

    # Actualizar DATA, FDATA, COSTOS, BREAKDOWN, FCLIENTES, etc. en el HTML
    actualizar_html(clientes_mes, total_mes, ing_usd, tc_mes,
                    horas_data, breakdown, resources, cres, costos, fdata, fegr_extra)

    # Actualizar sección HISTORIAL PASIVOS (hardcodeada en HTML)
    print("  Aplicando datos de HISTORIAL PASIVOS...")
    with open(OUTPUT_HTML, 'r', encoding='utf-8') as f: html_p = f.read()
    html_p = actualizar_pasivo_html(html_p, cierres, detalle, vivo)
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f: f.write(html_p)
    print(f"  OK guardado con pasivos ({len(html_p):,} bytes)")

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
    # Determinar último mes real con datos (antes de agregar meses futuros vacíos)
    ORDEN_M = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC']
    ultimo_real = 'AGO'
    for _m in reversed(ORDEN_M):
        if _m in fdata and (fdata[_m].get('ing_mx',0) or 0) > 0:
            ultimo_real = _m
            break
    html_fix = aplicar_fixes(html_fix, ultimo_mes_real=ultimo_real, fegr_extra=fegr_extra, tc_mes=tc_mes)
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f: f.write(html_fix)

    print()
    print("="*60)
    print("  LISTO — Recarga el dashboard.html en el navegador")
    print("="*60)

if __name__ == "__main__":
    main()
