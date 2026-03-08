# Pipeline Financiero — Comercial Andina S.A.C.

Pipeline de datos contable-financiero end-to-end para una empresa importadora de café ficticia. Construido como proyecto de portafolio integrando SQL Server, Python, n8n y Power BI.
## Stack

| Herramienta | Uso |
|---|---|
| SQL Server Express | Base de datos central |
| Python (pyodbc, openpyxl) | Generación de data y estados financieros |
| Power BI Desktop | Dashboard interactivo |
| n8n | Automatización del pipeline |
## Arquitectura
SQL Server → Python → Excel (SUNAT) → Power BI Dashboard
                                    ↓
                               n8n (Schedule semanal → Email)

## Base de datos

Motor: SQL Server Express (`LAPTOP-PCT25RMC\SQLEXPRESS`)  
Base de datos: `ComercialAndina`

### Tablas

| Tabla | Descripción |
|---|---|
| `plan_cuentas` | Plan contable PCGE — 35 cuentas |
| `centros_costo` | 6 centros: Ventas, Administración, Logística, Importaciones, Publicidad, Finanzas |
| `empleados` | 8 empleados con sueldo, régimen AFP/ONP y centro de costo |
| `asientos_cabecera` | Cabecera de asientos contables — periodo, fecha, glosa, tipo |
| `asientos_detalle` | Detalle de asientos — debe, haber, cuenta, centro |
| `tipo_cambio` | TC compra/venta mensual enero–junio 2026 |

### Volumen de data

- 6 meses: enero a junio 2026
- ~1,344 asientos contables generados
- Tipos: Venta, Costo de Ventas, Compra, Importación, Gasto Operativo, Planilla, Impuesto
## Scripts Python
     `generar_asientos.py`
Genera e inserta 6 meses de asientos contables directamente en SQL Server.

Incluye:
- 80 facturas de venta/mes con IGV
- 50 facturas de compra/mes
- 2 importaciones/mes con derechos aduaneros
- Gastos operativos fijos (alquiler, servicios, transporte)
- Planilla completa con cálculo real de 5ta categoría por tramos SUNAT
- Provisiones: CTS, gratificación, vacaciones
- Liquidación IGV mensual

    `estados_financieros.py`
Conecta a SQL Server, extrae saldos por periodo y genera un Excel con 4 hojas:

1. **Estado de Resultados** — Formato SUNAT 3.20
2. **Balance General** — Formato SUNAT 3.1 completo
3. **Balance de Comprobación** — debe, haber, saldos finales por cuenta
4. **Ratios Financieros** — liquidez corriente, prueba ácida, margen bruto, margen neto, ROA, rotación CxC

Uso:
```python
PERIODO = "2026-03"  # cambia el mes aquí
```

    `Function_5taCategoria.py`
Función de cálculo de retención de 5ta categoría por tramos según UIT vigente (S/ 5,500).

## Dashboard Power BI

Conectado directo a SQL Server. Incluye:

- **5 KPIs**: Ventas Netas, Costo de Ventas, Utilidad Bruta, Margen Neto %, CxC Pendiente
- **Barras**: Ventas vs Costos vs Utilidad Bruta por periodo
- **Dona**: Gastos por centro de costo
- **Cascada**: Flujo Ventas → Costo → Utilidad Bruta → Gastos → Utilidad Neta
- **Gauge**: Ventas acumuladas vs meta semestral
- **Slicer**: Filtro por periodo (enero–junio 2026)

### Medidas DAX principales
```dax
Ventas Netas = CALCULATE(SUM(asientos_detalle[haber]), asientos_cabecera[tipo_asiento] = "Venta")
Utilidad Bruta = [Ventas Netas] - [Costo de Ventas]
Margen Neto % = DIVIDE([Utilidad Bruta], [Ventas Netas], 0)

## Automatización n8n

Flujo semanal (lunes 8am):


Schedule Trigger → Code (pipeline summary) → Gmail (reporte por email)

## Empresa ficticia

**Comercial Andina S.A.C.**  
RUC: 20123456789  
Rubro: Importadora y distribuidora de café — Lima, Perú  
Empleados: 8 | Facturación semestral: ~S/ 1.54M

---

## Autor

**Bruno Portal Cossio**  
Estudiante de Contabilidad — ISIL, Lima Perú  
[LinkedIn](https://www.linkedin.com/in/bruno-portal-cossio) · [GitHub](https://github.com/brunoportal72)
