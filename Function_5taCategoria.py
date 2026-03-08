def calcular_5ta_categoria(sueldo_bruto_mensual):
    UIT = 5500
    anual = sueldo_bruto_mensual * 14
    renta_neta = max(anual - (7 * UIT), 0)

    if renta_neta <= 5 * UIT:
        impuesto = renta_neta * 0.08

    elif renta_neta <= 20 * UIT:
        impuesto = (5 * UIT * 0.08) + ((renta_neta - (5 * UIT)) * 0.14)

    elif renta_neta <= 35 * UIT:
        impuesto = (5 * UIT * 0.08) + (15 * UIT * 0.14) + ((renta_neta - (20 * UIT)) * 0.17)

    elif renta_neta <= 45 * UIT:
        impuesto = (5 * UIT * 0.08) + (15 * UIT * 0.14) + (15 * UIT * 0.17) + ((renta_neta - (35 * UIT)) * 0.20)

    else:
        impuesto = (5 * UIT * 0.08) + (15 * UIT * 0.14) + (15 * UIT * 0.17) + (10 * UIT * 0.20) + ((renta_neta - (45 * UIT)) * 0.30)

    return round(max(impuesto, 0) / 12, 2)
