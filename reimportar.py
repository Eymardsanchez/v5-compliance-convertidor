# -*- coding: utf-8 -*-
"""
Reimportación del Excel revisado → Pólizas COI (Etapa 2, rebanada 3).

Flujo
-----
1. Lee las hojas de banco del `.xlsm` revisado (columna A = ID estable).
2. Hace UPSERT por ID en Firebird: persiste las correcciones que el usuario hizo
   en el Excel (UUID pegado en col H, TIPO MATCH "Manual", etc.) SIN sobrescribir
   otras filas.
3. Reconstruye los objetos (ResultadoBanco/MovConciliado/MovBanco/CFDI) DESDE la
   BD — autosuficiente, no necesita los XML originales — para alimentar
   `ventana_polizas.VentanaPolizas` / `polizas_coi.generar_polizas`.

Columnas del reporte (tras poner el ID en la col A, todo corrido +1):
   A=1 ID · B=2 FECHA · C=3 FOLIO · D=4 DESC · E=5 RETIRO · F=6 DEPOSITO ·
   G=7 SALDO · H=8 UUID CFDI · I=9 UUID REP · Q=17 TIPO MATCH
"""
from __future__ import annotations

from types import SimpleNamespace

import openpyxl

from persistencia import (FirebirdPersistencia, FilaConciliacion, CFDIData,
                          ConflictoConcurrencia)
from conciliacion import MovBanco, MovConciliado
from cfdi_parser import CFDI

_C_ID, _C_FECHA, _C_DESC, _C_RET, _C_DEP, _C_UUID, _C_UREP, _C_TM = 1, 2, 4, 5, 6, 8, 9, 17


def _resumen_fila(f) -> str:
    """Texto corto de una fila para la bitácora de conflictos."""
    return (f"{getattr(f,'fecha','')} | {getattr(f,'concepto','')[:60]} | "
            f"dep={getattr(f,'deposito',0)} ret={getattr(f,'retiro',0)} | "
            f"uuid={getattr(f,'uuid_cfdi','')[:20]}")


def _resumen_mov(m) -> str:
    return (f"{getattr(m,'fecha','')} | {getattr(m,'descripcion','')[:60]} | "
            f"dep={getattr(m,'deposito',0)} ret={getattr(m,'retiro',0)} "
            f"saldo={getattr(m,'saldo',0)}")


def _es_hoja_banco(ws) -> bool:
    try:
        return (str(ws.cell(row=3, column=2).value or "").strip().upper() == "FECHA"
                and "UUID" in str(ws.cell(row=3, column=8).value or "").upper())
    except Exception:
        return False


def _num(v) -> float:
    try:
        return float(v) if v not in (None, "") else 0.0
    except Exception:
        return 0.0


def _fecha_txt(v) -> str:
    """Normaliza una celda de fecha a texto corto (<=10). Si Excel la entrego como
    datetime/date la formatea a ISO 'YYYY-MM-DD'; si es texto, recorta a 10. Evita
    el error Firebird -303 (FECHA es VARCHAR(10))."""
    if v in (None, ""):
        return ""
    try:
        from datetime import datetime as _dt, date as _date
        if isinstance(v, (_dt, _date)):
            return v.strftime("%Y-%m-%d")
    except Exception:
        pass
    return str(v).strip()[:10]


def leer_filas_excel(ruta: str) -> list[dict]:
    """Lee las filas de todas las hojas de banco del reporte revisado."""
    wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    filas: list[dict] = []
    for ws in wb.worksheets:
        if not _es_hoja_banco(ws):
            continue
        for row in ws.iter_rows(min_row=4):
            def cell(c):
                idx = c - 1
                return row[idx].value if idx < len(row) else None
            fecha = cell(_C_FECHA)
            if fecha in (None, ""):
                continue                      # fin de datos / pie de hoja
            fecha_txt = _fecha_txt(fecha)
            raw_uuid = str(cell(_C_UUID) or "").strip()
            # La celda de UUID puede traer varios: "AAA\n+ BBB\n+ CCC"
            partes = [p.strip() for p in raw_uuid.replace("\n+ ", "\n").split("\n")
                      if p.strip()]
            filas.append(dict(
                id=cell(_C_ID), banco=ws.title, fecha=fecha_txt,
                desc=str(cell(_C_DESC) or ""),
                retiro=_num(cell(_C_RET)), deposito=_num(cell(_C_DEP)),
                uuid=(partes[0] if partes else ""),
                uuid_suma=",".join(partes[1:]),
                uuid_rep=str(cell(_C_UREP) or "").strip(),
                tipo_match=str(cell(_C_TM) or "").strip()))
    wb.close()
    return filas


def _cfdi_de(cd: CFDIData) -> CFDI:
    """Construye un CFDI (cfdi_parser) desde los datos guardados en la BD."""
    c = CFDI()
    c.uuid            = cd.uuid
    c.tipo            = cd.tipo
    c.rfc_emisor      = cd.rfc_emisor
    c.rfc_receptor    = cd.rfc_receptor
    c.nombre_emisor   = cd.nombre_emisor
    c.nombre_receptor = cd.nombre_receptor
    c.subtotal        = cd.subtotal
    c.total           = cd.total
    c.descuento       = cd.descuento
    c.iva16           = cd.iva          # en la BD guardamos iva16 bajo 'iva'
    c.iva_retenido    = cd.iva_retenido
    c.isr_retenido    = cd.isr_retenido
    c.clave_prod_serv = cd.clave_prod_serv
    c.fecha_emision   = cd.fecha_emision
    return c


def subir_excel(ruta: str, rfc: str,
                persist: FirebirdPersistencia | None = None,
                resolver=None, usuario: str = "") -> dict:
    """SOLO SUBE: lee el .xlsm revisado y persiste las correcciones en Firebird
    (upsert por ID estable, sin sobrescribir otras filas). NO genera pólizas ni
    reconstruye nada — así el guardado y la impresión no se cruzan.
    Devuelve stats {filas, upsert, conflictos, descartados}."""
    propia = persist is None
    if propia:
        persist = FirebirdPersistencia()
        persist.inicializar()
    stats = dict(filas=0, upsert=0, conflictos=0, descartados=0, sin_id=0,
                 versionadas=0)
    try:
        filas_xl = leer_filas_excel(ruta)
        stats["filas"] = len(filas_xl)
        for fx in filas_xl:
            if not fx["id"]:
                continue
            try:
                fid = int(fx["id"])
            except Exception:
                continue
            tipo = "recibido" if fx["retiro"] else ("emitido" if fx["deposito"] else "")
            prev = persist.leer_fila(fid)
            if not prev:
                # El ID del Excel NO existe en la BD (Excel renumerado por la macro,
                # o no corresponde a esta conciliacion). NO se crea fila huerfana:
                # se omite y se reporta para que el usuario suba el Excel correcto.
                stats["sin_id"] += 1
                continue
            # ¿Cambio de MONTO o CONCEPTO? -> NO se sobrescribe la FILA original:
            # se registra en FILA_MODIFICACION (tabla aparte) con el valor nuevo, el
            # anterior y una observacion. La FILA conserva sus montos/concepto.
            _cambio_valor = (
                round(float(prev.retiro or 0), 2)   != round(float(fx["retiro"] or 0), 2) or
                round(float(prev.deposito or 0), 2) != round(float(fx["deposito"] or 0), 2) or
                (prev.concepto or "").strip()       != (fx["desc"] or "").strip())
            if _cambio_valor:
                try:
                    _mid = persist.registrar_modificacion(
                        origen_id=fid, rfc=rfc, banco=prev.banco, fecha=fx["fecha"],
                        concepto=fx["desc"], retiro=fx["retiro"], deposito=fx["deposito"],
                        concepto_ant=prev.concepto, retiro_ant=prev.retiro,
                        deposito_ant=prev.deposito,
                        observacion=("Monto/concepto modificado desde ID %d" % fid),
                        usuario=usuario)
                    stats["versionadas"] += 1
                except Exception:
                    pass
            # UPSERT en su lugar SOLO las correcciones; CONSERVA montos/concepto/fecha
            # originales de la BD (los cambios de valor ya quedaron en la bitacora).
            f = FilaConciliacion(
                id=fid, rfc=rfc, banco=prev.banco, fecha=prev.fecha,
                importe=prev.importe, concepto=prev.concepto,
                uuid_cfdi=fx["uuid"], uuid_rep=fx["uuid_rep"],
                tipo_match=fx["tipo_match"], tipo=(prev.tipo or tipo),
                uuid_suma=fx["uuid_suma"], cuenta=prev.cuenta, tercero=prev.tercero,
                deposito=prev.deposito, retiro=prev.retiro)
            f.conciliacion_id = prev.conciliacion_id
            ver_esp = prev.version_reg
            try:
                persist.upsert_fila(f, version_esperada=ver_esp, usuario=usuario)
                stats["upsert"] += 1
            except ConflictoConcurrencia as c:
                decision = (resolver(c) if resolver else "descartar")
                persist.registrar_conflicto(
                    "FILA", c.registro_id, decision, c.version_local,
                    c.version_guardada, _resumen_fila(c.local),
                    _resumen_fila(c.guardado), usuario=usuario)
                stats["conflictos"] += 1
                if decision == "sobrescribir":
                    actual = persist.leer_fila(fid)
                    persist.upsert_fila(
                        f, version_esperada=(actual.version_reg if actual else None),
                        usuario=usuario)
                    stats["upsert"] += 1
                else:
                    stats["descartados"] += 1
        return stats
    finally:
        if propia:
            persist.cerrar()


def construir_desde_bd(rfc: str, persist: FirebirdPersistencia | None = None,
                       conciliacion_ids=None):
    """SOLO REIMPRIME: reconstruye (resultados, recibidos, emitidos, stats) DESDE la
    BD para la empresa dada. No lee ningún Excel ni escribe nada.

    `conciliacion_ids`: si se da una lista, INTEGRA solo esas conciliaciones (uno o
    varios meses / todo el año, a elección del usuario); si es None, usa TODAS las
    filas del RFC."""
    propia = persist is None
    if propia:
        persist = FirebirdPersistencia()
        persist.inicializar()
    stats = dict(filas=0, conciliados=0, sin_datos_cfdi=0)
    try:
        por_banco: dict = {}
        from collections import Counter as _Counter
        from conciliacion import _mes_fecha as _mesf
        if conciliacion_ids:
            filas = []
            for _cid in conciliacion_ids:
                try:
                    filas.extend(persist.leer_filas(int(_cid)))
                except Exception:
                    pass
        else:
            filas = persist.leer_filas_por_rfc(rfc)
        stats["filas"] = len(filas)
        for _i, f in enumerate(filas):
            # Cada estado de cuenta = (conciliacion, banco): NO combinar el
            # mismo banco de conciliaciones/meses distintos en una sola hoja.
            _key = (getattr(f, "conciliacion_id", 0) or 0, f.banco)
            res = por_banco.get(_key)
            if res is None:
                res = SimpleNamespace(nombre=f.banco, conciliados=[], sin_cfdi=[],
                                      _banco=f.banco, _meses=_Counter())
                por_banco[_key] = res
            _mm = _mesf(f.fecha or "")
            if _mm:
                res._meses[_mm] += 1
            mb = MovBanco()
            mb.fecha = f.fecha; mb.descripcion = f.concepto
            mb.retiro = f.retiro; mb.deposito = f.deposito
            mb.id_fila = f.id
            mb.orden = _i      # orden de lectura (= ID = orden del PDF) para no reordenar
            if f.uuid_cfdi and f.tipo_match:
                cd = persist.obtener_cfdi(f.uuid_cfdi)
                if cd is None:
                    stats["sin_datos_cfdi"] += 1
                    res.sin_cfdi.append(mb)
                    continue
                suma = []
                for u in [x for x in f.uuid_suma.split(",") if x]:
                    cds = persist.obtener_cfdi(u)
                    if cds:
                        suma.append(_cfdi_de(cds))
                _cfdi_obj = _cfdi_de(cd)
                _tipo = f.tipo or ("recibido" if f.retiro else "emitido")
                # Restaurar los campos de DISPLAY (serie, folio CFDI, RFC, nombre,
                # fecha emisión, estatus); sin esto el reporte sale con esas
                # columnas vacías y el cruce a Recibidos/Emitidos queda incompleto.
                try:
                    from conciliacion import Conciliador as _C
                    _C._llenar_mov(mb, _cfdi_obj, tipo=_tipo)
                except Exception:
                    pass
                mb.uuid_rep  = f.uuid_rep or ""
                mb.uuid_suma = f.uuid_suma or ""
                mc = MovConciliado(
                    banco=mb, cfdi=_cfdi_obj, tipo=_tipo,
                    cfdi_suma=suma, revisar=bool(f.revisar),
                    tipo_match=f.tipo_match)
                res.conciliados.append(mc)
                stats["conciliados"] += 1
            else:
                res.sin_cfdi.append(mb)
        resultados = list(por_banco.values())
        # Nombre UNICO por estado de cuenta. Si el mismo banco aparece en
        # varias conciliaciones/meses, se distingue con el mes (y contador).
        _MESAB = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
                  7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}
        _rep = _Counter(r._banco for r in resultados)
        _usados: dict = {}
        for r in resultados:
            base = r._banco or "Banco"
            if _rep[r._banco] > 1:
                _mm = r._meses.most_common(1)[0][0] if r._meses else 0
                if _mm:
                    base = f"{r._banco} {_MESAB.get(_mm, '')}".strip()
            nombre = base[:31]
            _n = _usados.get(nombre, 0) + 1
            _usados[nombre] = _n
            r.nombre = nombre if _n == 1 else f"{base[:27]} ({_n})"

        rfc_u = (rfc or "").strip().upper()
        recibidos, emitidos = [], []
        for cd in persist.leer_todos_cfdis():
            cfdi = _cfdi_de(cd)
            if (cd.rfc_receptor or "").strip().upper() == rfc_u:
                recibidos.append(cfdi)
            elif (cd.rfc_emisor or "").strip().upper() == rfc_u:
                emitidos.append(cfdi)
        return resultados, recibidos, emitidos, stats
    finally:
        if propia:
            persist.cerrar()


def reconciliar_desde_bd(rfc: str, persist: FirebirdPersistencia | None = None,
                         meses=None, fases=None, grupos=None,
                         usuario: str = "", log=None,
                         recibidos=None, emitidos=None,
                         regenerar: bool = False) -> dict:
    """Re-concilia EN SU LUGAR las filas 'sin CFDI' del RFC contra los CFDIs YA
    guardados en la BD que aun no se han usado. NO crea filas nuevas ni toca las
    ya conciliadas; respeta TODAS las fases (usa `conciliar_multibanco`).

    `meses`: iterable de meses (1-12) a procesar; None = todos los que tengan banco.
    Devuelve stats {sin_cfdi_previas, reconciliadas, meses_objetivo, meses_con_banco,
    meses_sin_banco}."""
    from conciliacion import conciliar_multibanco as _cmb, _mes_fecha as _mf
    propia = persist is None
    if propia:
        persist = FirebirdPersistencia(); persist.inicializar()
    stats = dict(sin_cfdi_previas=0, reconciliadas=0, limpiadas=0,
                 meses_objetivo=[], meses_con_banco=[], meses_sin_banco=[])
    try:
        filas = persist.leer_filas_por_rfc(rfc)
        meses_con_banco = sorted({_mf(f.fecha) for f in filas if _mf(f.fecha)})
        stats["meses_con_banco"] = meses_con_banco
        objetivo = sorted({int(m) for m in meses}) if meses else list(meses_con_banco)
        stats["meses_objetivo"] = objetivo
        stats["meses_sin_banco"] = [m for m in objetivo if m not in meses_con_banco]
        meses_ok = set(objetivo) & set(meses_con_banco)
        if not meses_ok:
            return stats
        if regenerar:
            # REGENERAR: borrar el tipo_match (y el CFDI vinculado) de las filas
            # de los meses objetivo para re-conciliarlas DESDE CERO con los datos
            # actualizados. Asi cambios (p.ej. nombres) sí re-evaluan filas que
            # ya estaban marcadas.
            for f in filas:
                if _mf(f.fecha) in meses_ok and (f.uuid_cfdi or f.tipo_match):
                    f2 = FilaConciliacion(
                        id=f.id, rfc=rfc, banco=f.banco, fecha=f.fecha,
                        importe=f.importe, concepto=f.concepto,
                        uuid_cfdi="", uuid_rep="", tipo_match="", revisar=False,
                        tipo=f.tipo, uuid_suma="", cuenta=f.cuenta, tercero=f.tercero,
                        deposito=f.deposito, retiro=f.retiro)
                    f2.conciliacion_id = f.conciliacion_id
                    try:
                        persist.upsert_fila(f2, version_esperada=f.version_reg,
                                            usuario=usuario)
                        stats["limpiadas"] += 1
                    except Exception:
                        pass
            # Releer con los matches ya borrados (ahora todas 'sin CFDI').
            filas = persist.leer_filas_por_rfc(rfc)
        # CFDIs YA usados (no reutilizarlos): uuid_cfdi + uuid_suma de conciliadas.
        usados = set()
        for f in filas:
            if f.uuid_cfdi:
                usados.add(f.uuid_cfdi.strip().lower())
            for u in (f.uuid_suma or "").split(","):
                if u.strip():
                    usados.add(u.strip().lower())
        # Reconstruir MovBanco SOLO de las filas SIN CFDI de los meses objetivo.
        por_banco: dict = {}
        idx: dict = {}
        map_fila: dict = {}
        for f in filas:
            m = _mf(f.fecha)
            es_sin = not (f.uuid_cfdi and f.tipo_match)
            if not es_sin or m not in meses_ok:
                continue
            stats["sin_cfdi_previas"] += 1
            mb = MovBanco()
            mb.fecha = f.fecha; mb.descripcion = f.concepto
            mb.retiro = f.retiro; mb.deposito = f.deposito
            mb.mes = m; mb.id_fila = f.id
            i = idx.get(f.banco, 0); mb.orden = i; idx[f.banco] = i + 1
            por_banco.setdefault(f.banco, []).append(mb)
            map_fila[f.id] = f
        if not por_banco:
            return stats
        bancos_input = [(nombre, movs) for nombre, movs in por_banco.items()]
        rfc_u = (rfc or "").strip().upper()
        def _no_usado(c):
            return (getattr(c, "uuid", "") or "").strip().lower() not in usados
        if recibidos is not None or emitidos is not None:
            # CFDIs FRESCOS (del escaneo de XML): traen metodo/forma de pago -> alta
            # fidelidad para las fases (PUE/PPD). Solo los que no se han usado.
            rec_use = [c for c in (recibidos or []) if _no_usado(c)]
            emi_use = [c for c in (emitidos or []) if _no_usado(c)]
        else:
            # Fallback autosuficiente: reconstruye desde la tabla CFDI de la BD.
            rec_use, emi_use = [], []
            for cd in persist.leer_todos_cfdis():
                if (cd.uuid or "").strip().lower() in usados:
                    continue
                cfdi = _cfdi_de(cd)
                if (cd.rfc_receptor or "").strip().upper() == rfc_u:
                    rec_use.append(cfdi)
                elif (cd.rfc_emisor or "").strip().upper() == rfc_u:
                    emi_use.append(cfdi)
        resultados = _cmb(bancos_input, rec_use, emi_use,
                          grupos=grupos, fases=fases, log=log)
        # Actualizar EN SU LUGAR (por ID) las filas que pasaron a conciliadas.
        for res in resultados:
            for mc in getattr(res, "conciliados", []):
                mb = mc.banco
                fid = getattr(mb, "id_fila", None)
                if not fid or fid not in map_fila:
                    continue
                prev = map_fila[fid]
                _suma = ",".join(getattr(c, "uuid", "") for c in getattr(mc, "cfdi_suma", []) if getattr(c, "uuid", ""))
                f2 = FilaConciliacion(
                    id=fid, rfc=rfc, banco=prev.banco, fecha=prev.fecha,
                    importe=prev.importe, concepto=prev.concepto,
                    uuid_cfdi=(mb.uuid or ""), uuid_rep=(getattr(mb, "uuid_rep", "") or ""),
                    tipo_match=(mc.tipo_match or ""), revisar=bool(getattr(mc, "revisar", False)),
                    tipo=(mc.tipo or prev.tipo), uuid_suma=_suma,
                    cuenta=prev.cuenta, tercero=prev.tercero,
                    deposito=prev.deposito, retiro=prev.retiro)
                f2.conciliacion_id = prev.conciliacion_id
                try:
                    persist.upsert_fila(f2, version_esperada=prev.version_reg,
                                        usuario=usuario)
                    stats["reconciliadas"] += 1
                except Exception:
                    pass
        return stats
    finally:
        if propia:
            persist.cerrar()


# ── Reimportación del Excel del CONVERTIDOR (movimientos PDF→Excel) ───────────
# Layout: hoja "Movimientos", encabezado en fila 4. A=ID, B=FECHA, C=FOLIO,
# D=DESCRIPCION, E=DEPOSITO, F=RETIRO, G=SALDO.
_MC_ID, _MC_FECHA, _MC_FOLIO, _MC_DESC, _MC_DEP, _MC_RET, _MC_SALDO = 1, 2, 3, 4, 5, 6, 7


def reimportar_movimientos(ruta: str, persist=None, resolver=None,
                           usuario: str = "") -> dict:
    """Lee un Excel de movimientos ya modificado y actualiza la BD SOLO en lo que
    cambió (match por ID). Devuelve stats {leidos, nuevos, actualizados, sin_cambio}.
    Si otro proceso modificó un movimiento, se consulta `resolver` y se registra el
    conflicto en la bitácora de auditoría."""
    from persistencia import FirebirdPersistencia, MovimientoBanco
    propia = persist is None
    if propia:
        persist = FirebirdPersistencia()
        persist.inicializar()
    stats = dict(leidos=0, nuevos=0, actualizados=0, sin_cambio=0,
                 conflictos=0, descartados=0)
    try:
        wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
        ws = wb["Movimientos"] if "Movimientos" in wb.sheetnames else wb.worksheets[0]
        for row in ws.iter_rows(min_row=5):
            def cell(c):
                idx = c - 1
                return row[idx].value if idx < len(row) else None
            fecha = cell(_MC_FECHA)
            if fecha in (None, ""):
                continue
            _f = str(fecha).upper()
            if "SALDO INICIAL" in _f or "TOTAL" in _f:
                continue
            try:
                mid = int(cell(_MC_ID)) if cell(_MC_ID) not in (None, "") else 0
            except Exception:
                mid = 0
            m = MovimientoBanco(
                id=mid, fecha=str(fecha or ""), folio=str(cell(_MC_FOLIO) or ""),
                descripcion=str(cell(_MC_DESC) or ""),
                deposito=_num(cell(_MC_DEP)), retiro=_num(cell(_MC_RET)),
                saldo=_num(cell(_MC_SALDO)))
            # banco: conservar el de la BD si la fila ya existía
            prev = None
            if mid:
                prev = persist.leer_movimiento(mid)
                if prev:
                    m.banco = prev.banco
                    m.conversion_id = prev.conversion_id
            stats["leidos"] += 1
            ver_esp = prev.version_reg if prev else None
            try:
                r = persist.upsert_movimiento_si_cambio(m, version_esperada=ver_esp)
                _k = {"nuevo": "nuevos", "actualizado": "actualizados",
                      "sin_cambio": "sin_cambio"}.get(r, r)
                stats[_k] = stats.get(_k, 0) + 1
            except ConflictoConcurrencia as c:
                decision = (resolver(c) if resolver else "descartar")
                persist.registrar_conflicto(
                    "MOV_BANCO", c.registro_id, decision, c.version_local,
                    c.version_guardada, _resumen_mov(c.local),
                    _resumen_mov(c.guardado), usuario=usuario)
                stats["conflictos"] += 1
                if decision == "sobrescribir":
                    actual = persist.leer_movimiento(mid)
                    persist.upsert_movimiento_si_cambio(
                        m, version_esperada=(actual.version_reg if actual else None))
                    stats["actualizados"] += 1
                else:
                    stats["descartados"] += 1
        wb.close()
        return stats
    finally:
        if propia:
            persist.cerrar()
