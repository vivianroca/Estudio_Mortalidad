"""IND-010, 2009–2024: reproducir mes y día de ocurrencia en formato largo."""

import calendar
import csv
import hashlib
from collections import Counter
from pathlib import Path

import pandas as pd

from transformar_ind002 import entero, normalizar, unico
from transformar_ind004 import leer_hoja

RAIZ = Path(__file__).resolve().parents[1]
ANIOS = tuple(range(2009,2025))
MESES = ('Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto',
         'Septiembre','Octubre','Noviembre','Diciembre')
COLUMNAS = ['anio','mes_ocurrencia','dia_ocurrencia','conteo']
SALIDA = RAIZ / 'data/processed/ind010_2009_2024.csv'
REPORTE = RAIZ / 'reports/estructura/validacion_ind010_2009_2024.txt'


def transformar(anio,nombre,informe):
    ruta,huella,matriz,dimensiones = leer_hoja(anio,nombre)
    errores=[]
    informe.extend([f'\nAÑO: {anio}; archivo: {ruta.name}; hoja: {nombre}',
                    f'Dimensiones originales: {dimensiones}; SHA-256: {huella}.'])
    r_enc=unico([i for i,row in enumerate(matriz) if all(v in row for v in
                ('Día de ocurrencia','Total',*MESES))],'Fila de encabezados')
    h=matriz[r_enc]
    cd=unico([c for c,v in enumerate(h) if v=='Día de ocurrencia'],'Columna día')
    ct=unico([c for c,v in enumerate(h) if v=='Total'],'Columna total')
    cols={mes:unico([c for c,v in enumerate(h) if v==mes],mes) for mes in MESES}
    filas=[];totales=[];notas=[]
    for r in range(r_enc+1,len(matriz)):
        row=matriz[r]
        if not any(v!='' for v in row):
            continue
        etiqueta=row[cd]
        if isinstance(etiqueta,str) and normalizar(etiqueta).startswith(('nota:','notas:','fuente:')):
            notas.append(f'Fila {r+1}: {etiqueta}')
            if any(v!='' for c,v in enumerate(row) if c!=cd):
                errores.append(f'Fila {r+1}: nota con celdas adicionales')
            continue
        if any(v!='' for c,v in enumerate(row) if c not in {cd,ct,*cols.values()}):
            errores.append(f'Fila {r+1}: celdas fuera del cuadro')
        agregado=normalizar(etiqueta)=='todos los dias'
        try:
            dia=None if agregado else entero(etiqueta,f'fila {r+1}: día')
            if not agregado and dia not in range(1,32):
                raise ValueError(f'Fila {r+1}: día fuera del rango 1–31')
            valores={mes:entero(row[c],f'{anio}, fila {r+1}, {mes}') for mes,c in cols.items()}
            total=entero(row[ct],f'{anio}, fila {r+1}, Total')
        except ValueError as error:
            errores.append(str(error))
            continue
        f=dict(fila=r+1,dia=dia,valores=valores,total=total)
        (totales if agregado else filas).append(f)
    if len(totales)!=1:
        errores.append(f'Fila Todos los días: encontradas {len(totales)}')
    dias=Counter(f['dia'] for f in filas)
    if set(dias)!=set(range(1,32)) or any(n!=1 for n in dias.values()):
        errores.append('Días faltantes, repetidos o inesperados')
    informe.extend([f'Encabezados: fila {r_enc+1}; meses originales: {cols}.',
                    'Todos los días y columna Total se excluyen del CSV y se usan para validar.',
                    'Se conservan los 31 días de cada mes, incluidos los ceros originales de fechas inexistentes; no representan fechas válidas.',
                    'No se imputan guiones, vacíos ni no numéricos.',
                    'Notas/fuentes: '+('; '.join(notas) or 'ninguna en la hoja.')])
    if errores:
        informe.extend('ERROR: '+e for e in errores)
        return None
    nacional=totales[0]
    for f in [nacional,*filas]:
        suma=sum(f['valores'].values())
        informe.append(f'Fila {f["fila"]}: Total={f["total"]}; suma de meses={suma}; diferencia={suma-f["total"]}.')
        if suma!=f['total']:
            errores.append(f'Fila {f["fila"]}: suma de meses no coincide con Total')
    for mes in MESES:
        suma=sum(f['valores'][mes] for f in filas)
        publicado=nacional['valores'][mes]
        informe.append(f'{mes}: total original={publicado}; suma de días={suma}; diferencia={suma-publicado}.')
        if suma!=publicado:
            errores.append(f'{mes}: días no suman el total mensual')
    invalidas=[]
    for numero,mes in enumerate(MESES,1):
        limite=calendar.monthrange(anio,numero)[1]
        for f in filas:
            if f['dia']>limite:
                invalidas.append((mes,f['dia'],f['valores'][mes]))
                if f['valores'][mes]!=0:
                    errores.append(f'Fecha inexistente con conteo no cero: {mes}, día {f["dia"]}, valor {f["valores"][mes]}')
    informe.append(f'Fechas inexistentes revisadas contra calendario y año bisiesto: {invalidas!r}.')
    ancho=pd.DataFrame([{'dia_ocurrencia':f['dia'],**f['valores']} for f in filas])
    largo=ancho.melt(id_vars=['dia_ocurrencia'],value_vars=list(MESES),
                      var_name='mes_ocurrencia',value_name='conteo')
    largo.insert(0,'anio',anio)
    largo=largo[COLUMNAS].astype({'anio':'int64','dia_ocurrencia':'int64','conteo':'int64'})
    fuente={(mes,f['dia']):f['valores'][mes] for f in filas for mes in MESES}
    actual={(r.mes_ocurrencia,r.dia_ocurrencia):r.conteo for r in largo.itertuples(index=False)}
    if fuente!=actual or len(largo)!=372 or largo.duplicated(COLUMNAS[:-1]).any():
        errores.append('Intersecciones incorrectas o claves duplicadas después de melt')
    total=int(largo.conteo.sum())
    informe.append(f'TOTAL GENERAL: publicado={nacional["total"]}; transformado={total}; diferencia={total-nacional["total"]}.')
    if total!=nacional['total'] or sum(f['total'] for f in filas)!=nacional['total']:
        errores.append('Total general no coincide con datos o márgenes')
    if hashlib.sha256(ruta.read_bytes()).hexdigest()!=huella:
        errores.append('El hash del Excel original cambió durante la ejecución')
    informe.extend([f'Intersecciones verificadas: {len(largo)}.',
                    'Guiones, vacíos y no numéricos en datos y márgenes: 0.',
                    'EJEMPLOS:',largo.head(5).to_csv(index=False).strip()])
    informe.extend('ERROR: '+e for e in errores)
    informe.append(f'RESULTADO {anio}: '+('NO VALIDADO.' if errores else 'VALIDADO.'))
    print(f'{anio}: {len(largo)} registros; total={total}; '+('NO VALIDADO' if errores else 'VALIDADO'))
    for error in errores:
        print('  '+error)
    return None if errores else largo


def main():
    with (RAIZ/'reports/estructura/catalogo_indicadores.csv').open(encoding='utf-8-sig') as f:
        indicador=unico([r for r in csv.DictReader(f) if r['id_provisional']=='IND-010'],'IND-010')
    informe=['VALIDACIÓN IND-010 — 2009–2024',
             'Formato largo: '+', '.join(COLUMNAS),
             'Se reproducen todas las celdas analíticas de la matriz 31×12, sin transformar los Excel originales.',
             'Las combinaciones de fecha inexistentes se conservan únicamente como ceros publicados; no imputados.']
    resultados=[]
    for anio in ANIOS:
        try:
            datos=transformar(anio,indicador[str(anio)],informe)
            if datos is not None:
                resultados.append(datos)
        except Exception as error:
            mensaje=f'{anio}: {type(error).__name__}: {error}'
            informe.append('ERROR: '+mensaje);print(mensaje)
    exito=len(resultados)==len(ANIOS)
    if exito:
        datos=pd.concat(resultados,ignore_index=True)
        orden={mes:i for i,mes in enumerate(MESES,1)}
        datos=datos.sort_values(['anio','mes_ocurrencia','dia_ocurrencia'],
                               key=lambda c:c.map(orden) if c.name=='mes_ocurrencia' else c,
                               kind='stable').reset_index(drop=True)
        SALIDA.parent.mkdir(parents=True,exist_ok=True)
        temporal=SALIDA.with_suffix('.tmp.csv')
        try:
            datos.to_csv(temporal,index=False,encoding='utf-8-sig')
            pd.testing.assert_frame_equal(datos,pd.read_csv(temporal,encoding='utf-8-sig'))
            temporal.replace(SALIDA)
            informe.append(f'RESULTADO GLOBAL: VALIDADO; {len(datos)} registros; {SALIDA}')
        except Exception as error:
            exito=False
            informe.append(f'ERROR exportando: {type(error).__name__}: {error}')
        finally:
            temporal.unlink(missing_ok=True)
    if not exito:
        informe.append(f'RESULTADO GLOBAL: NO VALIDADO; {len(resultados)}/16 años válidos. CSV no generado ni actualizado.')
        if SALIDA.exists():
            informe.append('Un archivo anterior no queda validado por esta ejecución.')
    REPORTE.parent.mkdir(parents=True,exist_ok=True)
    REPORTE.write_text('\n'.join(informe)+'\n',encoding='utf-8')
    print(informe[-1]);print(f'Reporte: {REPORTE}')
    if not exito:
        raise SystemExit(1)


if __name__=='__main__':
    main()
