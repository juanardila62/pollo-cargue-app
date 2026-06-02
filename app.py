from flask import Flask, render_template, request, jsonify, send_file
import openpyxl
from datetime import datetime, date
import os, copy, json, tempfile

app = Flask(__name__)

# ── Ruta del Excel: usa variable de entorno en producción ─────────────────────
DATA_DIR   = os.environ.get('DATA_DIR', os.path.join(os.path.dirname(__file__), 'data'))
EXCEL_NAME = 'CONSOLIDADO INFORMES CARGUES DE POLLO 2026.xlsx'
EXCEL_PATH = os.path.join(DATA_DIR, EXCEL_NAME)
CONFIG_FILE= os.path.join(DATA_DIR, 'config.json')

os.makedirs(DATA_DIR, exist_ok=True)

# Si no hay Excel en DATA_DIR pero sí localmente, copiarlo automáticamente
_LOCAL_EXCEL = '/Users/macair/Documents/SAVICOL/' + EXCEL_NAME
if not os.path.exists(EXCEL_PATH) and os.path.exists(_LOCAL_EXCEL):
    import shutil
    shutil.copy2(_LOCAL_EXCEL, EXCEL_PATH)

COL_DIA=2;COL_FECHA=3;COL_VIAJE=4;COL_GRANJA=5;COL_CUAD=6
COL_PERS=7;COL_COND=8;COL_PLACA=9;COL_HINICIO=10;COL_HSALIDA=11
COL_CANT=12;COL_TOTAL=13
SHEET_NAME='CONSOLIDADO 2026'
DATA_START=6

# ── Config ────────────────────────────────────────────────────────────────────
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE) as f: return json.load(f)
    return {}

def save_config(cfg):
    with open(CONFIG_FILE,'w') as f: json.dump(cfg,f)

# ── Excel helpers ─────────────────────────────────────────────────────────────
def load_wb():
    return openpyxl.load_workbook(EXCEL_PATH)

def fmt_time(t):
    if t is None: return ''
    if isinstance(t,datetime): return t.strftime('%H:%M')
    try: return f'{t.hour:02d}:{t.minute:02d}'
    except: return str(t)

def get_entries():
    if not os.path.exists(EXCEL_PATH): return []
    wb=load_wb(); ws=wb[SHEET_NAME]; rows=[]
    for r in ws.iter_rows(min_row=DATA_START, values_only=True):
        if all(v is None for v in r): continue
        fecha=r[COL_FECHA-1]
        if fecha is None: continue
        fecha_s=fecha.strftime('%Y-%m-%d') if isinstance(fecha,(datetime,date)) else str(fecha)
        rows.append({'dia':r[COL_DIA-1],'fecha':fecha_s,'viaje':r[COL_VIAJE-1],
                     'granja':r[COL_GRANJA-1] or '','cuadrilla':r[COL_CUAD-1] or '',
                     'personas':r[COL_PERS-1],'conductor':r[COL_COND-1] or '',
                     'placa':r[COL_PLACA-1] or '','hora_inicio':fmt_time(r[COL_HINICIO-1]),
                     'hora_salida':fmt_time(r[COL_HSALIDA-1]),'cantidad':r[COL_CANT-1],
                     'total_dia':r[COL_TOTAL-1]})
    return rows

def next_viaje(fecha_s):
    vs=[e['viaje'] for e in get_entries() if e['fecha']==fecha_s and e['viaje'] is not None]
    return (max(vs)+1) if vs else 1

def last_dia():
    if not os.path.exists(EXCEL_PATH): return 1
    wb=load_wb(); ws=wb[SHEET_NAME]; d=1
    for r in ws.iter_rows(min_row=DATA_START, values_only=True):
        v=r[COL_DIA-1]
        if v is not None and isinstance(v,(int,float)): d=int(v)
    return d

def parse_time(s):
    if not s: return None
    from datetime import time as dtime
    for fmt in ('%H:%M','%H:%M:%S'):
        try:
            dt=datetime.strptime(s.strip(),fmt); return dtime(dt.hour,dt.minute,dt.second)
        except: pass
    return None

def write_row(ws, nr, lr, fecha_dt, viaje, dia_num, data):
    def sc(col,val):
        cell=ws.cell(row=nr,column=col)
        src=ws.cell(row=lr,column=col)
        if src.has_style: cell._style=copy.copy(src._style)
        cell.value=val
    sc(COL_DIA,dia_num); sc(COL_FECHA,fecha_dt); sc(COL_VIAJE,viaje)
    sc(COL_GRANJA,(data.get('granja','') or '').upper())
    sc(COL_CUAD,(data.get('cuadrilla','CARGUEROS') or 'CARGUEROS').upper())
    sc(COL_PERS,int(data.get('personas',11) or 11))
    sc(COL_COND,(data.get('conductor','') or '').upper())
    sc(COL_PLACA,(data.get('placa','') or '').upper())
    sc(COL_HINICIO,parse_time(data.get('hora_inicio','')))
    sc(COL_HSALIDA,parse_time(data.get('hora_salida','')))
    sc(COL_CANT,int(data['cantidad']) if data.get('cantidad') else None)
    sc(COL_TOTAL,int(data['total_dia']) if data.get('total_dia') else None)

# ── Prompt ────────────────────────────────────────────────────────────────────
PROMPT = """Eres un asistente que extrae datos de reportes de cargue de pollos en Colombia.

REGLAS:
1. PLACAS colombianas: exactamente 3 letras + 3 números (ej: TLO881, EQY779).
   Whisper las transcribe mal: "EQY es 181"→EQY181, "EQY setenta y ocho uno"→EQY781.
2. HORAS: "once treinta y cinco"=11:35, "tres y veintidós"=03:22. Formato HH:MM 24h.
3. NÚMEROS: "tres mil quinientos sesenta y cuatro"=3564.
4. NOMBRES: nombres propios colombianos.

Para cada viaje devuelve:
- conductor: nombre en MAYÚSCULAS
- placa: 3 letras + 3 números, sin espacios
- hora_inicio: HH:MM 24h
- hora_salida: HH:MM 24h
- cantidad: entero
- total_dia: entero si se menciona como total del día, sino null
- granja: MALAGUEÑA/GARCERAS/SAN JOSE/HUERTAS/etc. o null
- cuadrilla: "CARGUEROS" por defecto
- personas: 11 por defecto

Responde SOLO con JSON válido:
{"viajes": [...]}

Transcripción:
"""

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html', today=date.today().strftime('%Y-%m-%d'))

@app.route('/api/config', methods=['GET','POST'])
def api_config():
    if request.method=='POST':
        cfg=load_config(); cfg.update(request.json); save_config(cfg)
        return jsonify({'ok':True})
    cfg=load_config(); cfg.pop('api_key',None)  # no exponer la clave
    return jsonify(cfg)

@app.route('/api/entries')
def api_entries():
    return jsonify(get_entries())

@app.route('/api/next-viaje')
def api_next_viaje():
    fecha=request.args.get('fecha',date.today().strftime('%Y-%m-%d'))
    return jsonify({'next_viaje':next_viaje(fecha)})

@app.route('/api/has-key')
def api_has_key():
    cfg=load_config()
    return jsonify({'has_key': bool(cfg.get('api_key','').strip())})

@app.route('/api/transcribe', methods=['POST'])
def api_transcribe():
    if 'audio' not in request.files:
        return jsonify({'ok':False,'error':'No se recibió archivo'}),400
    f=request.files['audio']
    api_key=load_config().get('api_key','')
    if not api_key:
        return jsonify({'ok':False,'error':'Configura la clave API primero'}),400
    suffix=os.path.splitext(f.filename)[1] or '.ogg'
    with tempfile.NamedTemporaryFile(suffix=suffix,delete=False) as tmp:
        f.save(tmp.name); tmp_path=tmp.name
    try:
        from groq import Groq
        client=Groq(api_key=api_key)
        with open(tmp_path,'rb') as audio_file:
            result=client.audio.transcriptions.create(
                file=(f.filename, audio_file),
                model='whisper-large-v3',
                language='es',
                response_format='text'
            )
        texto=result if isinstance(result,str) else result.text
        return jsonify({'ok':True,'texto':texto.strip()})
    except Exception as e:
        return jsonify({'ok':False,'error':str(e)}),500
    finally:
        os.unlink(tmp_path)

@app.route('/api/parse', methods=['POST'])
def api_parse():
    try:
        body=request.json
        texto=body.get('texto','').strip()
        api_key=body.get('api_key') or load_config().get('api_key','')
        fecha=body.get('fecha',date.today().strftime('%Y-%m-%d'))
        proveedor=load_config().get('proveedor','groq')
        if not api_key: return jsonify({'ok':False,'error':'Clave API no configurada'}),400
        if not texto:   return jsonify({'ok':False,'error':'Texto vacío'}),400

        full=PROMPT+texto
        if proveedor=='anthropic':
            import anthropic
            c=anthropic.Anthropic(api_key=api_key)
            msg=c.messages.create(model='claude-haiku-4-5-20251001',max_tokens=2048,
                                   messages=[{'role':'user','content':full}])
            raw=msg.content[0].text.strip()
        else:
            from groq import Groq
            c=Groq(api_key=api_key)
            chat=c.chat.completions.create(model='llama-3.3-70b-versatile',
                messages=[{'role':'user','content':full}],max_tokens=2048,temperature=0.1)
            raw=chat.choices[0].message.content.strip()

        if raw.startswith('```'): raw=raw.split('\n',1)[1].rsplit('```',1)[0].strip()
        viajes=json.loads(raw).get('viajes',[])
        v=next_viaje(fecha)
        for i,x in enumerate(viajes): x['viaje_num']=v+i
        return jsonify({'ok':True,'viajes':viajes})
    except Exception as e:
        return jsonify({'ok':False,'error':str(e)}),500

@app.route('/api/add-bulk', methods=['POST'])
def api_add_bulk():
    body=request.json; viajes=body.get('viajes',[]); fecha_s=body.get('fecha',date.today().strftime('%Y-%m-%d'))
    if not viajes: return jsonify({'ok':False,'error':'Sin viajes'}),400
    try:
        wb=load_wb(); ws=wb[SHEET_NAME]
        fecha_dt=datetime.strptime(fecha_s,'%Y-%m-%d')
        entries=get_entries()
        fe=[e for e in entries if e['fecha']==fecha_s]
        dia_num=fe[0]['dia'] if fe and fe[0]['dia'] else last_dia()
        last_row=ws.max_row
        for r in range(ws.max_row,DATA_START-1,-1):
            if any(ws.cell(row=r,column=c).value is not None for c in range(1,14)):
                last_row=r; break
        vs=next_viaje(fecha_s)
        for i,v in enumerate(viajes):
            write_row(ws,last_row+1+i,last_row,fecha_dt,vs+i,dia_num,v)
        wb.save(EXCEL_PATH)
        return jsonify({'ok':True,'guardados':len(viajes)})
    except Exception as e:
        return jsonify({'ok':False,'error':str(e)}),500

@app.route('/api/add', methods=['POST'])
def api_add():
    data=request.json
    try:
        wb=load_wb(); ws=wb[SHEET_NAME]
        fecha_s=data.get('fecha',date.today().strftime('%Y-%m-%d'))
        fecha_dt=datetime.strptime(fecha_s,'%Y-%m-%d')
        viaje=data.get('viaje') or next_viaje(fecha_s)
        entries=get_entries()
        fe=[e for e in entries if e['fecha']==fecha_s]
        dia_num=fe[0]['dia'] if fe and fe[0]['dia'] else last_dia()
        last_row=ws.max_row
        for r in range(ws.max_row,DATA_START-1,-1):
            if any(ws.cell(row=r,column=c).value is not None for c in range(1,14)):
                last_row=r; break
        write_row(ws,last_row+1,last_row,fecha_dt,viaje,dia_num,data)
        wb.save(EXCEL_PATH)
        return jsonify({'ok':True,'viaje':viaje})
    except Exception as e:
        return jsonify({'ok':False,'error':str(e)}),500

@app.route('/api/delete-last', methods=['POST'])
def api_delete_last():
    try:
        wb=load_wb(); ws=wb[SHEET_NAME]
        for r in range(ws.max_row,DATA_START-1,-1):
            if any(ws.cell(row=r,column=c).value is not None for c in range(1,14)):
                ws.delete_rows(r); wb.save(EXCEL_PATH)
                return jsonify({'ok':True})
        return jsonify({'ok':False,'error':'No hay filas'})
    except Exception as e:
        return jsonify({'ok':False,'error':str(e)}),500

@app.route('/api/download-excel')
def api_download_excel():
    if not os.path.exists(EXCEL_PATH):
        return jsonify({'error':'No hay archivo Excel'}),404
    return send_file(EXCEL_PATH, as_attachment=True, download_name=EXCEL_NAME)

@app.route('/api/upload-excel', methods=['POST'])
def api_upload_excel():
    if 'excel' not in request.files:
        return jsonify({'ok':False,'error':'No se recibió archivo'}),400
    f=request.files['excel']
    f.save(EXCEL_PATH)
    return jsonify({'ok':True})

@app.route('/api/open-folder', methods=['POST'])
def api_open_folder():
    import subprocess
    subprocess.Popen(['open', DATA_DIR])
    return jsonify({'ok':True})

if __name__=='__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT',5050)))
