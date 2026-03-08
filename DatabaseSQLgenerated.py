import pyodbc
import random
from Function_5taCategoria import calcular_5ta_categoria
from datetime import date, timedelta

conexion = pyodbc.connect(
    "DRIVER={SQL Server};"
    "SERVER=LAPTOP-PCT25RMC\\SQLEXPRESS;"
    "DATABASE=ComercialAndina;"
    "Trusted_Connection=yes;"
)
cursor = conexion.cursor()

CENTROS = {
    "Ventas":              1,
    "Administración":      2,
    "Logística y Almacén": 3,
    "Importaciones":       4,
    "Publicidad":          5,
    "Finanzas":            6,
}

CTA = {
    "caja":              "10.1",
    "bcp":               "10.41.1",
    "interbank":         "10.41.2",
    "cxc":               "12.13",
    "igv_credito":       "16.73",
    "mercaderias":       "20.111",
    "igv_ventas":        "40.111",
    "ir_3ra":            "40.171",
    "ir_5ta":            "40.173",
    "essalud_pagar":     "40.31",
    "onp_pagar":         "40.32",
    "sueldos_pagar":     "41.11",
    "grati_pagar":       "41.14",
    "vacac_pagar":       "41.15",
    "cts_pagar":         "41.51",
    "afp_pagar":         "41.7",
    "cxp":               "42.12",
    "capital":           "50.1",
    "utilidades":        "59.11",
    "ventas":            "70.121",
    "costo_ventas":      "69.121",
    "compras":           "60.11",
    "derechos_aduana":   "60.913",
    "sueldos_gasto":     "62.11",
    "grati_gasto":       "62.14",
    "vacac_gasto":       "62.15",
    "essalud_gasto":     "62.71",
    "pensiones_gasto":   "62.72",
    "cts_gasto":         "62.91",
    "renta_5ta_gasto":   "627",
    "alquiler":          "63.52",
    "luz":               "63.61",
    "agua":              "63.63",
    "internet":          "63.65",
    "transporte":        "63.111",
}

EMPLEADOS = [
    {"nombre": "Carlos Mendoza", "centro": "Finanzas",            "sueldo": 6000, "regimen": "AFP"},
    {"nombre": "Ana Quispe",     "centro": "Administración",      "sueldo": 4500, "regimen": "AFP"},
    {"nombre": "Luis Paredes",   "centro": "Ventas",              "sueldo": 3800, "regimen": "AFP"},
    {"nombre": "María Torres",   "centro": "Ventas",              "sueldo": 3200, "regimen": "ONP"},
    {"nombre": "Jorge Huanca",   "centro": "Logística y Almacén", "sueldo": 2800, "regimen": "ONP"},
    {"nombre": "Rosa Ccallo",    "centro": "Logística y Almacén", "sueldo": 2500, "regimen": "AFP"},
    {"nombre": "Diego Vargas",   "centro": "Importaciones",       "sueldo": 4200, "regimen": "AFP"},
    {"nombre": "Lucía Sánchez",  "centro": "Publicidad",          "sueldo": 2200, "regimen": "ONP"},
]

MESES = [
    ("2026-01", date(2026, 1, 1),  date(2026, 1, 31)),
    ("2026-02", date(2026, 2, 1),  date(2026, 2, 28)),
    ("2026-03", date(2026, 3, 1),  date(2026, 3, 31)),
    ("2026-04", date(2026, 4, 1),  date(2026, 4, 30)),
    ("2026-05", date(2026, 5, 1),  date(2026, 5, 31)),
    ("2026-06", date(2026, 6, 1),  date(2026, 6, 30)),
]

IGV_TASA      = 0.18
VENTAS_MES    = 80
COMPRAS_MES   = 50
TICKET_VENTA  = (1500, 4000)
TICKET_COMPRA = (800, 2500)

def fecha_aleatoria(inicio, fin):
    delta = (fin - inicio).days
    return inicio + timedelta(days=random.randint(0, delta))

def redondear(valor):
    return round(valor, 2)

def insertar_cabecera(periodo, fecha, glosa, tipo_asiento):
    cursor.execute("""
        INSERT INTO asientos_cabecera (periodo, fecha, glosa, tipo_asiento)
        VALUES (?, ?, ?, ?);
        SELECT @@IDENTITY AS id;
    """, (periodo, str(fecha), glosa, tipo_asiento))
    cursor.nextset()
    return int(cursor.fetchone()[0])

def insertar_detalle(id_asiento, cuenta, centro_id, debe, haber):
    cursor.execute("""
        INSERT INTO asientos_detalle (id_asiento, codigo_cuenta, id_centro, debe, haber)
        VALUES (?, ?, ?, ?, ?)
    """, id_asiento, cuenta, centro_id, debe, haber)

def generar_ventas(periodo, inicio, fin):
    for i in range(VENTAS_MES):
        fecha  = fecha_aleatoria(inicio, fin)
        base   = redondear(random.uniform(*TICKET_VENTA))
        igv    = redondear(base * IGV_TASA)
        total  = redondear(base + igv)
        glosa  = f"Venta café - Factura {periodo}-V{i+1:03d}"

        id_cab = insertar_cabecera(periodo, fecha, glosa, "Venta")
        insertar_detalle(id_cab, CTA["cxc"],        CENTROS["Ventas"], total, 0)
        insertar_detalle(id_cab, CTA["igv_ventas"], CENTROS["Ventas"], 0, igv)
        insertar_detalle(id_cab, CTA["ventas"],     CENTROS["Ventas"], 0, base)

        costo    = redondear(base * 0.60)
        id_costo = insertar_cabecera(periodo, fecha,
                    f"Costo venta café - {periodo}-V{i+1:03d}", "Costo de Ventas")
        insertar_detalle(id_costo, CTA["costo_ventas"], CENTROS["Ventas"],              costo, 0)
        insertar_detalle(id_costo, CTA["mercaderias"],  CENTROS["Logística y Almacén"], 0, costo)

def generar_compras(periodo, inicio, fin):
    for i in range(COMPRAS_MES):
        fecha  = fecha_aleatoria(inicio, fin)
        base   = redondear(random.uniform(*TICKET_COMPRA))
        igv    = redondear(base * IGV_TASA)
        total  = redondear(base + igv)
        glosa  = f"Compra mercadería café - {periodo}-C{i+1:03d}"

        id_cab = insertar_cabecera(periodo, fecha, glosa, "Compra")
        insertar_detalle(id_cab, CTA["compras"],     CENTROS["Logística y Almacén"], base, 0)
        insertar_detalle(id_cab, CTA["igv_credito"], CENTROS["Logística y Almacén"], igv,  0)
        insertar_detalle(id_cab, CTA["cxp"],         CENTROS["Importaciones"],       0, total)

def generar_importacion(periodo, inicio, fin):
    for i in range(2):
        fecha           = fecha_aleatoria(inicio, fin)
        base            = redondear(random.uniform(8000, 20000))
        igv             = redondear(base * IGV_TASA)
        arancel         = redondear(base * 0.06)
        total_proveedor = redondear(base + igv)
        glosa           = f"Importación café - {periodo}-IMP{i+1:02d}"

        id_cab = insertar_cabecera(periodo, fecha, glosa, "Importación")
        insertar_detalle(id_cab, CTA["compras"],         CENTROS["Importaciones"], base,            0)
        insertar_detalle(id_cab, CTA["igv_credito"],     CENTROS["Importaciones"], igv,             0)
        insertar_detalle(id_cab, CTA["derechos_aduana"], CENTROS["Importaciones"], arancel,         0)
        insertar_detalle(id_cab, CTA["cxp"],             CENTROS["Importaciones"], 0, total_proveedor)
        insertar_detalle(id_cab, CTA["cxp"],             CENTROS["Importaciones"], 0, arancel)

def generar_gastos_operativos(periodo, inicio, fin):
    gastos = [
        (CTA["alquiler"],   "Administración",      4500, "Alquiler local almacén"),
        (CTA["luz"],        "Administración",        380, "Energía eléctrica"),
        (CTA["agua"],       "Administración",        120, "Agua"),
        (CTA["internet"],   "Administración",        180, "Internet y telefonía"),
        (CTA["transporte"], "Logística y Almacén",   800, "Flete distribución"),
        (CTA["transporte"], "Importaciones",        1200, "Flete importación"),
    ]
    for cuenta, centro_nombre, monto, descripcion in gastos:
        fecha  = fecha_aleatoria(inicio, fin)
        igv    = redondear(monto * IGV_TASA)
        total  = redondear(monto + igv)
        glosa  = f"{descripcion} - {periodo}"

        id_cab = insertar_cabecera(periodo, fecha, glosa, "Gasto Operativo")
        insertar_detalle(id_cab, cuenta,             CENTROS[centro_nombre], monto, 0)
        insertar_detalle(id_cab, CTA["igv_credito"], CENTROS[centro_nombre], igv,   0)
        insertar_detalle(id_cab, CTA["cxp"],         CENTROS[centro_nombre], 0, total)

def generar_planilla(periodo, inicio, fin):
    fecha         = fin
    total_bruto   = 0
    total_essalud = 0
    total_onp     = 0
    total_afp     = 0
    total_5ta     = 0
    total_grati   = 0
    total_vacac   = 0
    total_cts     = 0
    total_neto    = 0

    for emp in EMPLEADOS:
        s         = emp["sueldo"]
        essalud   = redondear(s * 0.09)
        onp       = redondear(s * 0.13) if emp["regimen"] == "ONP" else 0
        afp       = redondear(s * 0.125) if emp["regimen"] == "AFP" else 0
        renta_5ta = calcular_5ta_categoria(s)
        grati     = redondear(s / 6)
        vacac     = redondear(s / 12)
        cts       = redondear(s / 12)
        neto      = redondear(s - onp - afp - renta_5ta)

        total_bruto   += s
        total_essalud += essalud
        total_onp     += onp
        total_afp     += afp
        total_5ta     += renta_5ta
        total_grati   += grati
        total_vacac   += vacac
        total_cts     += cts
        total_neto    += neto

    id_cab = insertar_cabecera(periodo, fecha, f"Planilla remuneraciones {periodo}", "Planilla")
    insertar_detalle(id_cab, CTA["sueldos_gasto"], CENTROS["Administración"], total_bruto, 0)
    insertar_detalle(id_cab, CTA["sueldos_pagar"], CENTROS["Administración"], 0, total_neto)
    insertar_detalle(id_cab, CTA["onp_pagar"],     CENTROS["Administración"], 0, total_onp)
    insertar_detalle(id_cab, CTA["afp_pagar"],     CENTROS["Administración"], 0, total_afp)
    insertar_detalle(id_cab, CTA["ir_5ta"],        CENTROS["Administración"], 0, total_5ta)

    id_cab2 = insertar_cabecera(periodo, fecha, f"Cargas sociales empleador {periodo}", "Planilla")
    insertar_detalle(id_cab2, CTA["essalud_gasto"], CENTROS["Administración"], total_essalud, 0)
    insertar_detalle(id_cab2, CTA["essalud_pagar"], CENTROS["Administración"], 0, total_essalud)

    id_cab3 = insertar_cabecera(periodo, fecha, f"Provisión gratificación {periodo}", "Planilla")
    insertar_detalle(id_cab3, CTA["grati_gasto"], CENTROS["Administración"], total_grati, 0)
    insertar_detalle(id_cab3, CTA["grati_pagar"], CENTROS["Administración"], 0, total_grati)

    id_cab4 = insertar_cabecera(periodo, fecha, f"Provisión vacaciones {periodo}", "Planilla")
    insertar_detalle(id_cab4, CTA["vacac_gasto"], CENTROS["Administración"], total_vacac, 0)
    insertar_detalle(id_cab4, CTA["vacac_pagar"], CENTROS["Administración"], 0, total_vacac)

    id_cab5 = insertar_cabecera(periodo, fecha, f"Provisión CTS {periodo}", "Planilla")
    insertar_detalle(id_cab5, CTA["cts_gasto"], CENTROS["Administración"], total_cts, 0)
    insertar_detalle(id_cab5, CTA["cts_pagar"], CENTROS["Administración"], 0, total_cts)

def generar_igv_mensual(periodo, fin):
    fecha       = fin
    igv_ventas  = redondear(random.uniform(18000, 28000))
    igv_compras = redondear(igv_ventas * 0.55)
    igv_neto    = redondear(igv_ventas - igv_compras)

    id_cab = insertar_cabecera(periodo, fecha, f"Liquidación IGV {periodo}", "Impuesto")
    insertar_detalle(id_cab, CTA["igv_ventas"],  CENTROS["Finanzas"], igv_ventas, 0)
    insertar_detalle(id_cab, CTA["igv_credito"], CENTROS["Finanzas"], 0, igv_compras)
    insertar_detalle(id_cab, CTA["ir_3ra"],      CENTROS["Finanzas"], 0, igv_neto)

print("Iniciando generación de asientos...")

for periodo, inicio, fin in MESES:
    print(f"  Generando {periodo}...")
    generar_ventas(periodo, inicio, fin)
    generar_compras(periodo, inicio, fin)
    generar_importacion(periodo, inicio, fin)
    generar_gastos_operativos(periodo, inicio, fin)
    generar_planilla(periodo, inicio, fin)
    generar_igv_mensual(periodo, fin)

conexion.commit()
cursor.close()
conexion.close()

print("✅ Todos los asientos generados e insertados correctamente.")