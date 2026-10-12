#!/usr/bin/env python3
"""
Genera el portal de liquidaciones de comisiones de Grupo MG.

Uso:
    python3 build_portal.py RESUMEN_COMISIONES.xlsx [--codigos CODIGOS_DE_ACCESO.xlsx]

Lee las hojas mensuales del RESUMEN COMISIONES, arma un archivo cifrado por persona
(carpeta d/) y escribe CODIGOS_DE_ACCESO.xlsx (PRIVADO: NO se sube a la web).
Los codigos ya entregados se conservan entre ejecuciones (se leen de CODIGOS_DE_ACCESO.xlsx).
"""
import sys, os, re, json, hashlib, secrets, unicodedata, datetime, glob, argparse
import openpyxl, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tools'))
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ITER = 200_000
ALFABETO = "23456789ABCDEFGHJKMNPQRSTUVWXYZ"
JEFES = ["LIZBETH CUELLAR", "CARLOS CASTILLO", "GUIDO COSSIO", "DANTE RUIZ"]
# Cuentas con vista total (completar con los nombres reales si se desea)
EXTERNOS = {"DIGITAL","GONZALES DE PENA","MARTINEZ VALDEZ ROBERT ELLIOT","BARBIER FLORES JUAN JOSE"}  # corredores/empresas sin acceso propio
GERENCIA = [("GERENTE COMERCIAL", "Gerencia Comercial"), ("GERENTE GENERAL", "Gerencia General")]
# Equipo vigente segun hoja CATALOGOS de la Revision de Ventas (se usa si la hoja del mes no trae bloque SUM)
EQUIPOS = {
 "CARLOS CASTILLO": ["DIANA DEL AGUILA","JORGE ALVARADO","DEYVIS BENDEZU","RENZO SOSA","SONIA CONTRERAS","BRENDA CARPIO"],
 "DANTE RUIZ": ["ANGELLO CARRION","HERNAN PEREZ","LUIS TAFUR","DANIEL ACEVEDO","RONALD DE LA CRUZ","RICARDO OLAYA","NATALIA QUINTANA"],
 "GUIDO COSSIO": ["ROBERTO SANTISTEBAN","MILAGROS DE LA CRUZ","JOSE ESTRADA","JOSE LUIS SANCHEZ","JEAN VARA","MIGUEL ZAMUDIO","OSCAR MARATARA"],
 "LIZBETH CUELLAR": ["CESAR ZAVALETA","ANDREA ARANGURI","GUSTAVO CHUCHON","RICARDO ALLCCA"],
}
# Escala de jefes (vigente desde 01/09/2025): unidades del equipo -> total por unidad
ESCALA_JEFES = [(1,6,400),(7,13,600),(14,18,800),(19,10**6,1000)]
MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii","ignore").decode().upper()
    return re.sub(r"\s+"," ",s).strip()

def titulo(s):
    return " ".join(w.capitalize() if w not in ("DE","LA","DEL") else w.lower() for w in norm(s).split())

def num(v):
    return v if isinstance(v,(int,float)) and not isinstance(v,bool) else None

def fecha_mes(v):
    return v.strftime("%Y-%m") if isinstance(v,(datetime.datetime,datetime.date)) else (str(v) if v else None)

def leer_resumen(path):
    wf = openpyxl.load_workbook(path)
    wv = openpyxl.load_workbook(path, data_only=True)
    meses = {}
    avisos = []
    for ws in wv.worksheets:
        if ws.sheet_state != "visible" or not ws.title.upper().startswith("RESUMEN"): continue
        wsf = wf[ws.title]
        filas, mes_key = [], None
        for r in range(3, min(ws.max_row, 80)+1):
            nombre = ws.cell(r,2).value
            if mes_key is None and isinstance(ws.cell(r,3).value,(datetime.datetime,datetime.date)):
                mes_key = ws.cell(r,3).value.strftime("%Y-%m")
            if not nombre: continue
            monto, ventas, desem, bonos, pen = [num(ws.cell(r,c).value) for c in (4,5,6,7,8)]
            coment = ws.cell(r,10).value
            if all(x is None for x in (monto,ventas,desem,bonos,pen)) and not coment: continue
            for c in (4,5):
                f = wsf.cell(r,c).value
                if isinstance(f,str) and f.startswith("=") and ws.cell(r,c).value is None:
                    avisos.append(f"{ws.title} fila {r}: formula sin valor calculado")
            pres = ws.cell(r,11).value
            filas.append(dict(
                fila=r, nombre=norm(nombre), monto=monto, ventas=ventas, desembolsos=desem, bonos=bonos, penalidad=pen,
                comentario=str(coment).strip() if coment else "",
                presentado=fecha_mes(pres) if pres else "",
                aprobado=str(ws.cell(r,12).value or "").strip().upper(),
                modalidad=str(ws.cell(r,13).value or "").strip().upper(),
                formula_e=wsf.cell(r,5).value if isinstance(wsf.cell(r,5).value,str) else None))
        if not mes_key: continue
        # bloques de equipo por formula =SUM(E4:E7)...
        for f in filas:
            f["jefe"] = ""
        for j in filas:
            if j["nombre"] in JEFES and j["formula_e"]:
                refs = set()
                for a,b in re.findall(r"E(\d+):E(\d+)", j["formula_e"]): refs.update(range(int(a),int(b)+1))
                for a in re.findall(r"(?<![:\d])E(\d+)(?![:\d])", j["formula_e"].replace("SUM(","(")):
                    refs.add(int(a))
                refs.discard(j["fila"])
                for f in filas:
                    if f["fila"] in refs and f["nombre"] not in JEFES: f["jefe"] = j["nombre"]
        for f in filas:
            if not f["jefe"] and f["nombre"] not in JEFES:
                for jj,eq in EQUIPOS.items():
                    if f["nombre"] in eq: f["jefe"] = jj
        meses[mes_key] = filas
    return meses, avisos

def total_bloque(b):
    for k in ("pagar","comision","total_items"):
        if b.get(k) is not None: return b[k]
    return sum(i["monto"] or 0 for i in b["items"])

def casar_usuario(nombre_bloque, usuarios):
    if not nombre_bloque: return None
    t = set(nombre_bloque.split())
    c = [u for u in usuarios if (set(u.split()) <= t or t <= set(u.split())) and (len(set(u.split()) & t) >= 2 or len(u.split()) == 1 or len(t) == 1)]
    if len(c) == 1: return c[0]
    # error de tipeo frecuente en nombres: MARAPARA / MARATARA
    t2 = {w.replace("MARAPARA","MARATARA") for w in t}
    c = [u for u in usuarios if set(u.split()) <= t2 or t2 <= set(u.split())]
    return c[0] if len(c) == 1 else None

def abrir_cuadro(path, clave):
    from leer_cuadro import leer_cuadro
    import msoffcrypto
    with open(path, "rb") as fh:
        of = msoffcrypto.OfficeFile(fh)
        if of.is_encrypted():
            tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False); tmp.close()
            of.load_key(password=clave)
            with open(tmp.name, "wb") as out: of.decrypt(out)
            path = tmp.name
    bl = leer_cuadro(path)
    return bl

def adjuntar_detalle(meses, bloques):
    """Une el detalle del Cuadro de Ventas a cada fila del Resumen. Devuelve lista de diferencias para revisar."""
    usuarios = {f["nombre"] for fs in meses.values() for f in fs}
    por_user = {}
    for b in bloques:
        u = casar_usuario(b["nombre"], usuarios)
        if u: por_user.setdefault(u, []).append(b)
    dif = []
    usados = set()
    for mk in sorted(meses):
        for f in meses[mk]:
            bs = por_user.get(f["nombre"], [])
            f["detalle"] = None
            if f["nombre"] in JEFES:
                cand = [b for b in bs if b["mes"] == mk and id(b) not in usados]
                if cand and total_bloque(cand[0]) is not None:
                    b = cand[0]; usados.add(id(b)); f["detalle"] = b
                    if f["monto"] is None: f["monto"] = total_bloque(b); f["monto_real"] = True
                continue
            if f["monto"] is None: continue
            cand = [b for b in bs if id(b) not in usados and abs(total_bloque(b) - f["monto"]) < 0.01]
            cand.sort(key=lambda b: b["mes"] != mk)
            if cand and (cand[0]["mes"] == mk or abs(int(cand[0]["mes"][:4])*12+int(cand[0]["mes"][5:]) - int(mk[:4])*12 - int(mk[5:])) <= 1):
                usados.add(id(cand[0])); f["detalle"] = cand[0]
            else:
                mism = [b for b in bs if b["mes"] == mk]
                dif.append((mk, titulo(f["nombre"]), f["monto"], [total_bloque(b) for b in mism] or "sin bloque en el Cuadro"))
    return dif

def estado(f):
    if f["aprobado"] in ("OK","APROBADO"): return "Aprobado"
    if f["presentado"]: return "Presentado a gerencia"
    return "En revisión"

def tramo_jefe(u):
    for a,b,t in ESCALA_JEFES:
        if a <= u <= b: return t
    return None

def fila_publica(f, mes):
    es_jefe = f["nombre"] in JEFES
    monto, ref = f["monto"], False
    if es_jefe and monto is None and f["ventas"] and mes >= "2025-09" and not f.get("detalle"):
        t = tramo_jefe(int(f["ventas"]))
        if t: monto, ref = f["ventas"]*t, True
    det = f.get("detalle")
    if es_jefe and f.get("monto_real"): ref = False
    pago = det["estado"] if det and det.get("estado") else ""
    return dict(detalle=[{k:i[k] for k in ("proyecto","dpto","cliente","tipo","concepto","minuta","escritura","monto","pendiente","nota")} for i in det["items"]] if det else None,
                bono=det["bono"] if det else None, dscto=det["dscto"] if det else None, pago=pago.title(),
                nombre=titulo(f["nombre"]), rol="Jefe de ventas" if es_jefe else "Asesor", jefe=titulo(f["jefe"]) if f["jefe"] else "",
                modalidad=f["modalidad"].title() if f["modalidad"] else "", monto=monto, referencial=ref,
                ventas=f["ventas"], desembolsos=f["desembolsos"], bonos=f["bonos"], penalidad=f["penalidad"],
                estado=estado(f), presentado=f["presentado"], comentario=f["comentario"])

def codigo_nuevo():
    s = "".join(secrets.choice(ALFABETO) for _ in range(10))
    return s[:5]+"-"+s[5:]

def limpiar(c): return re.sub(r"[^A-Z0-9]","",c.upper())

def archivo_de(c): return hashlib.sha256(("mgfile:"+limpiar(c)).encode()).hexdigest()[:32]+".enc"

def cifrar(obj, c):
    salt, iv = secrets.token_bytes(16), secrets.token_bytes(12)
    key = PBKDF2HMAC(hashes.SHA256(), 32, salt, ITER).derive(limpiar(c).encode())
    ct = AESGCM(key).encrypt(iv, json.dumps(obj, ensure_ascii=False, separators=(",",":")).encode(), None)
    return salt+iv+ct

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("resumen"); ap.add_argument("--codigos", default="CODIGOS_DE_ACCESO.xlsx")
    ap.add_argument("--salida", default="d"); ap.add_argument("--cuadro"); ap.add_argument("--clave", default=os.environ.get("CUADRO_CLAVE",""))
    a = ap.parse_args()
    meses, avisos = leer_resumen(a.resumen)
    for v in avisos: print("AVISO:", v)
    keys = sorted(meses)
    if a.cuadro:
        dif = adjuntar_detalle(meses, abrir_cuadro(a.cuadro, a.clave))
        wbd = openpyxl.Workbook(); wd = wbd.active; wd.title = "Diferencias"
        wd.append(["Mes","Persona","Monto en Resumen","Total en Cuadro de Ventas (o motivo)"])
        for r in dif: wd.append([r[0], r[1], r[2], str(r[3])])
        for c,w in zip("ABCD",(10,34,18,40)): wd.column_dimensions[c].width = w
        wbd.save("REPORTE_CONCILIACION.xlsx")
        print(f"Detalle del Cuadro adjuntado. {len(dif)} casos donde el Resumen no coincide con el Cuadro (ver REPORTE_CONCILIACION.xlsx; en esos casos no se muestra detalle).")
    # usuarios
    por_user = {}
    for mk in keys:
        for f in meses[mk]:
            por_user.setdefault(f["nombre"], True)
    modal = {}
    for mk in keys:
        for f in meses[mk]: modal.setdefault(f["nombre"], set()).add(f["modalidad"])
    def externo(n):
        return n.startswith("CORREDOR") or n in EXTERNOS or (modal[n] == {"FACTURA"})
    excluidos = sorted(n for n in por_user if externo(n))
    usuarios = [(n, "jefe" if n in JEFES else "asesor") for n in sorted(por_user) if not externo(n)]
    print("Sin acceso propio (corredores/empresas; sí aparecen en la vista de Gerencia):", ", ".join(excluidos))
    usuarios += [(norm(n), "gerencia") for n,_ in GERENCIA]
    etiquetas = dict((norm(n), l) for n,l in GERENCIA)
    # codigos previos
    previos = {}
    if os.path.exists(a.codigos):
        ws = openpyxl.load_workbook(a.codigos).active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and row[3]: previos[norm(row[0])] = row[3]
    for f in glob.glob(os.path.join(a.salida,"*.enc")): os.remove(f)
    os.makedirs(a.salida, exist_ok=True)
    filas_codigos = []
    for nombre, rol in usuarios:
        cod = previos.get(nombre) or codigo_nuevo()
        meses_u = []
        for mk in keys:
            filas = meses[mk]
            if rol == "gerencia":
                sel = filas
            elif rol == "jefe":
                sel = [f for f in filas if f["nombre"]==nombre or f["jefe"]==nombre]
            else:
                sel = [f for f in filas if f["nombre"]==nombre]
            if not sel: continue
            y, m = mk.split("-")
            meses_u.append(dict(clave=mk, etiqueta=f"{MESES[int(m)-1]} {y}", filas=[fila_publica(f, mk) for f in sel]))
        if not meses_u: continue
        payload = dict(usuario=dict(nombre=etiquetas.get(nombre) or titulo(nombre), rol=rol), generado=datetime.date.today().isoformat(), meses=meses_u)
        open(os.path.join(a.salida, archivo_de(cod)),"wb").write(cifrar(payload, cod))
        filas_codigos.append((etiquetas.get(nombre) or titulo(nombre), {"asesor":"Asesor","jefe":"Jefe de ventas","gerencia":"Gerencia"}[rol], len(meses_u), cod))
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Codigos"
    ws.append(["Nombre","Perfil","Meses con datos","Código de acceso","Enviado (marcar)"])
    for r in filas_codigos: ws.append(list(r))
    for c,w in zip("ABCDE",(28,16,16,20,16)): ws.column_dimensions[c].width = w
    wb.save(a.codigos)
    print(f"{len(filas_codigos)} accesos generados · {len(keys)} meses ({keys[0]} a {keys[-1]}) · códigos en {a.codigos}")

if __name__ == "__main__": main()
