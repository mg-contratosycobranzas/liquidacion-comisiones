"""Lee el CUADRO DE VENTAS - COMISIONES (una hoja por asesor, bloques mensuales) y devuelve el detalle por persona y mes."""
import re, datetime, unicodedata, openpyxl

def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii","ignore").decode().upper()
    return re.sub(r"\s+"," ",s).strip()

def iso(v):
    return v.strftime("%Y-%m-%d") if isinstance(v,(datetime.datetime,datetime.date)) else (str(v).strip() if v not in (None,"") else "")

def es_num(v): return isinstance(v,(int,float)) and not isinstance(v,bool)

def leer_hoja(ws, desde="2025-06"):
    rows = [list(r)+[None]*(20-len(r)) for r in ws.iter_rows(max_col=20, values_only=True)]
    bloques = []
    n = len(rows)
    i = 0
    while i < n:
        r = rows[i]
        if isinstance(r[1], str) and r[1].strip().upper() == "DPTO":
            # cabecera con nombre y mes: hasta 3 filas arriba
            nombre, mes = None, None
            for k in range(1,4):
                if i-k < 0: break
                p = rows[i-k]
                fechas = [c for c in p[4:7] if isinstance(c,(datetime.datetime,datetime.date))]
                if fechas:
                    mes = fechas[0].strftime("%Y-%m")
                    nombre = next((norm(c) for c in p[1:5] if isinstance(c,str) and len(c.strip())>2 and c.strip()!="|"), None)
                    break
            # columnas por encabezado (dos filas)
            h = {}
            for rr in (rows[i], rows[i+1] if i+1<n else []):
                for j,c in enumerate(rr):
                    if isinstance(c,str):
                        t = norm(c)
                        h.setdefault(t, j)
            col = dict(
                dpto=1, proyecto=h.get("PROYECTO",3), cliente=h.get("CLIENTE",4),
                minuta=h.get("MINUTA",5), escritura=h.get("ESCRITURA",6),
                tipo=h.get("FINANC.", h.get("TIPO")), monto=h.get("COMISION S/.", h.get("MONTO")),
                pend=h.get("PENDIENTE"), concepto=h.get("CONCEPTO"))
            # si MONTO apunta a la columna del encabezado 'MONTO' pero la comision esta en la fila siguiente, ya cubierto
            j = i+2
            items = []
            total_fila = None
            while j < n:
                q = rows[j]
                if isinstance(q[1], str) and "TOTAL POR PAGAR" in q[1].upper(): break
                if isinstance(q[1], str) and q[1].strip().upper() == "DPTO": break
                cm = q[col["monto"]] if col["monto"] is not None else None
                tiene_item = bool(str(q[col["cliente"]] or "").strip() or str(q[col["proyecto"]] or "").strip()) and (es_num(cm) or q[col["concepto"] or 0])
                if tiene_item:
                    extra = [str(c).strip() for c in q[(col["concepto"] or 9)+1:] if isinstance(c,str) and c.strip()]
                    items.append(dict(
                        dpto=str(q[1]).strip() if q[1] not in (None,"") else "", proyecto=str(q[col["proyecto"]] or "").strip(),
                        cliente=re.sub(r"\s+"," ",str(q[col["cliente"]] or "")).strip(),
                        minuta=iso(q[col["minuta"]]), escritura=iso(q[col["escritura"]]),
                        tipo=str(q[col["tipo"]] or "").strip() if col["tipo"] is not None else "",
                        monto=cm if es_num(cm) else None,
                        pendiente=q[col["pend"]] if col["pend"] is not None and es_num(q[col["pend"]]) else None,
                        concepto=str(q[col["concepto"]] or "").strip() if col["concepto"] is not None else "",
                        nota=" · ".join(extra)))
                elif es_num(cm) and not str(q[col["proyecto"]] or "").strip():
                    total_fila = cm
                j += 1
            # total por pagar
            tot = dict(comision=None,bono=None,dscto=None,pagar=None,estado="")
            if j < n:
                hdr = rows[j+1] if j+1<n else []
                val = rows[j+2] if j+2<n else []
                pos = {}
                bonos = []
                for k,c in enumerate(hdr):
                    if isinstance(c,str):
                        t = norm(c)
                        if t.startswith("COMISION"): pos["comision"]=k
                        elif t.startswith("BONO"): bonos.append(k)
                        elif t.startswith(("DSCTO","ADELANT","PRESTAMO")): pos["dscto"]=k
                        elif t.startswith("T.PAGAR"): pos["pagar"]=k
                if "comision" in pos and es_num(val[pos["comision"]]): tot["comision"]=val[pos["comision"]]
                tot["bono"] = sum(val[k] for k in bonos if es_num(val[k])) if bonos else None
                if "dscto" in pos and es_num(val[pos["dscto"]]): tot["dscto"]=val[pos["dscto"]]
                if "pagar" in pos and es_num(val[pos["pagar"]]): tot["pagar"]=val[pos["pagar"]]
                for c in val:
                    if isinstance(c,str) and norm(c) in ("PAGADO","PROGRAMADO","PENDIENTE","APROBADO"): tot["estado"]=norm(c)
            if mes and mes >= desde:
                bloques.append(dict(hoja=ws.title, nombre=nombre, mes=mes, items=items, total_items=total_fila, **tot))
            i = j
        i += 1
    return bloques

def leer_cuadro(path, desde="2025-06"):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible" or ws.title.upper() in ("IMPRIMIR",): continue
        try: out += leer_hoja(ws, desde)
        except Exception as e: print("AVISO hoja", ws.title, e)
    return out
