"""IND-009, 2009–2024: lugares de ocurrencia por residencia en formato largo."""

import csv
import hashlib
from collections import Counter
from pathlib import Path

import pandas as pd

from transformar_ind002 import DEPARTAMENTOS, entero, normalizar, unico
from transformar_ind004 import leer_hoja

RAIZ = Path(__file__).resolve().parents[1]
ANIOS = tuple(range(2009,2025))
SALIDA = RAIZ / 'data/processed/ind009_2009_2024.csv'
REPORTE = RAIZ / 'reports/estructura/validacion_ind009_2009_2024.txt'
COLUMNAS = ['anio','depto_residencia','municipio_residencia','lugar_ocurrencia','conteo']
ESPERADOS = {normalizar(d) for d in DEPARTAMENTOS}


def transformar(anio,nombre,informe):
    ruta,huella,matriz,dimensiones=leer_hoja(anio,nombre)
    errores=[]
    informe.extend([f'\nAÑO: {anio}; archivo: {ruta.name}; hoja: {nombre}',
                    f'Dimensiones originales: {dimensiones}; SHA-256: {huella}.'])
    r_enc=unico([i for i,row in enumerate(matriz) if all(e in row for e in
                ('Departamento de residencia','Municipio','Total','Lugar de ocurrencia'))], 'Encabezado principal')
    h=matriz[r_enc]
    cd,cm,ct=(h.index(e) for e in ('Departamento de residencia','Municipio','Total'))
    r_sub=r_enc+1
    sub=matriz[r_sub]
    lugares={c:v for c,v in enumerate(sub) if isinstance(v,str) and v.strip()
             and c not in (cd,cm,ct)}
    if not lugares or any(c<h.index('Lugar de ocurrencia') for c in lugares):
        raise ValueError('No se identificaron inequívocamente las columnas de lugar de ocurrencia')
    if len({normalizar(v) for v in lugares.values()})!=len(lugares):
        raise ValueError('Categorías de lugar de ocurrencia duplicadas')
    informe.append(f'Encabezados: filas {r_enc+1} y {r_sub+1}; lugares: {lugares}. Columnas indexadas desde cero.')
    filas=[];notas=[]
    for r in range(r_sub+1,len(matriz)):
        row=matriz[r]
        if not any(v!='' for v in row):
            continue
        d,m=row[cd],row[cm]
        if isinstance(d,str) and normalizar(d).startswith(('fuente:','nota:','notas:')):
            notas.append(f'Fila {r+1}: {d}')
            if any(v!='' for c,v in enumerate(row) if c!=cd):
                errores.append(f'Fila {r+1}: nota con celdas adicionales')
            continue
        if not isinstance(d,str) or not d.strip() or not isinstance(m,str) or not m.strip():
            errores.append(f'Fila {r+1}: departamento o municipio vacío/no textual')
            continue
        if any(v!='' for c,v in enumerate(row) if c not in {cd,cm,ct,*lugares}):
            errores.append(f'Fila {r+1}: celdas inesperadas fuera del cuadro')
        values={}
        for c in [ct,*lugares]:
            try:
                values[c]=entero(row[c],f'{anio}, fila {r+1}, columna {c+1}')
            except ValueError as error:
                errores.append(str(error))
        filas.append(dict(fila=r+1,depto=d,municipio=m,valores=values))
    if errores:
        informe.extend('ERROR: '+e for e in errores)
        return None,None
    nacionales=[f for f in filas if normalizar(f['depto']) in {'todos los departamentos','total pais','total'}]
    nacional=unico(nacionales,'Fila nacional')
    subtotales=[f for f in filas if f is not nacional and normalizar(f['municipio'])=='todos los municipios']
    analiticas=[f for f in filas if f is not nacional and normalizar(f['municipio'])!='todos los municipios']
    deps=Counter(normalizar(f['depto']) for f in subtotales)
    if set(deps)!=ESPERADOS or any(n!=1 for n in deps.values()):
        errores.append(f'Subtotales departamentales faltantes, inesperados o duplicados: {deps!r}')
    allowed=ESPERADOS|{'extranjero','ignorado'}
    for f in analiticas:
        if normalizar(f['depto']) not in allowed:
            errores.append(f'Fila {f["fila"]}: departamento desconocido')
    especiales=Counter(normalizar(f['depto']) for f in analiticas if normalizar(f['depto']) in {'extranjero','ignorado'})
    if especiales!={'extranjero':1,'ignorado':1}:
        errores.append(f'Categorías especiales faltantes/duplicadas: {especiales!r}')
    keys=[(normalizar(f['depto']),normalizar(f['municipio'])) for f in analiticas]
    if len(keys)!=len(set(keys)):
        errores.append('Municipios repetidos dentro del mismo departamento')
    informe.extend([f'Total nacional: fila {nacional["fila"]}, etiqueta {nacional["depto"]!r}.',
                    f'Subtotales departamentales: {len(subtotales)}; filas analíticas de residencia: {len(analiticas)}.',
                    'Totales nacionales, subtotales Todos los municipios y columna Total se excluyen del CSV.',
                    'Extranjero e Ignorado se conservan como categorías originales de residencia.',
                    'No se exige igual cantidad de municipios entre años ni se infieren ceros para categorías no publicadas.',
                    'Notas/fuentes: '+('; '.join(notas) or 'ninguna dentro de la hoja.')])
    for f in filas:
        calculado=sum(f['valores'][c] for c in lugares)
        if calculado!=f['valores'][ct]:
            errores.append(f'Fila {f["fila"]}: suma de lugares={calculado}, Total={f["valores"][ct]}, diferencia={calculado-f["valores"][ct]}')
    informe.append(f'Sumas de lugares comprobadas contra Total en {len(filas)} filas originales.')
    for f in subtotales:
        d=f['depto']
        partes=[p for p in analiticas if normalizar(p['depto'])==normalizar(d)]
        if not partes:
            errores.append(f'{d}: sin municipios analíticos')
        for c in [ct,*lugares]:
            calculado=sum(p['valores'][c] for p in partes)
            publicado=f['valores'][c]
            informe.append(f'Departamento {d}/{"Total" if c==ct else lugares[c]}: publicado={publicado}; municipios={calculado}; diferencia={calculado-publicado}.')
            if calculado!=publicado:
                errores.append(f'Municipios no suman subtotal: {d}, columna {c+1}, diferencia={calculado-publicado}')
    for c in [ct,*lugares]:
        calculado=sum(f['valores'][c] for f in analiticas)
        publicado=nacional['valores'][c]
        informe.append(f'Nacional/{"Total" if c==ct else lugares[c]}: publicado={publicado}; transformado={calculado}; diferencia={calculado-publicado}.')
        if calculado!=publicado:
            errores.append(f'Residencias no suman nacional: columna {c+1}, diferencia={calculado-publicado}')
        por_margenes=sum(f['valores'][c] for f in subtotales)+sum(f['valores'][c] for f in analiticas if normalizar(f['depto']) in {'extranjero','ignorado'})
        if por_margenes!=publicado:
            errores.append(f'Subtotales y especiales no suman nacional: columna {c+1}')
    ancho=pd.DataFrame([{'_fila':f['fila'],'depto_residencia':f['depto'],
                         'municipio_residencia':f['municipio'],**{v:f['valores'][c] for c,v in lugares.items()}}
                        for f in analiticas])
    largo=ancho.melt(id_vars=['_fila','depto_residencia','municipio_residencia'],
                      value_vars=list(lugares.values()),var_name='lugar_ocurrencia',value_name='conteo')
    largo.insert(0,'anio',anio)
    source={(f['fila'],v):(f['depto'],f['municipio'],f['valores'][c]) for f in analiticas for c,v in lugares.items()}
    actual={(r._0,r.lugar_ocurrencia):(r.depto_residencia,r.municipio_residencia,r.conteo)
            for r in largo[['_fila','lugar_ocurrencia','depto_residencia','municipio_residencia','conteo']].itertuples(index=False)}
    if source!=actual or len(largo)!=len(source) or largo.duplicated(COLUMNAS[:-1]).any():
        errores.append('Intersecciones incorrectas o claves duplicadas después de melt')
    if int(largo.conteo.sum())!=nacional['valores'][ct]:
        errores.append('El CSV no coincide con total general publicado')
    if hashlib.sha256(ruta.read_bytes()).hexdigest()!=huella:
        errores.append('El hash del Excel original cambió durante la ejecución')
    informe.extend([f'Intersecciones verificadas: {len(largo)}.',
                    'Guiones, vacíos y no numéricos en datos y márgenes: 0.',
                    'EJEMPLOS:',largo[COLUMNAS].head(5).to_csv(index=False).strip()])
    informe.extend('ERROR: '+e for e in errores)
    informe.append(f'RESULTADO {anio}: '+('NO VALIDADO.' if errores else 'VALIDADO.'))
    print(f'{anio}: {len(analiticas)} residencias × {len(lugares)} lugares = {len(largo)} registros; '+('NO VALIDADO' if errores else 'VALIDADO'))
    for error in errores:
        print('  '+error)
    estructura=dict(lugares=list(lugares.values()),residencias=len(analiticas),dimensiones=dimensiones)
    return (None,None) if errores else (largo[COLUMNAS],estructura)


def main():
    with (RAIZ/'reports/estructura/catalogo_indicadores.csv').open(encoding='utf-8-sig') as f:
        indicador=unico([r for r in csv.DictReader(f) if r['id_provisional']=='IND-009'],'IND-009')
    informe=['VALIDACIÓN IND-009 — 2009–2024',
             'Formato largo: '+', '.join(COLUMNAS),
             'Se conservan nombres originales; no se fusionan categorías que cambian entre años.',
             'No se imputan guiones, vacíos ni no numéricos. Una validación fallida bloquea la exportación conjunta.']
    resultados=[];estructuras={}
    for anio in ANIOS:
        try:
            datos,estructura=transformar(anio,indicador[str(anio)],informe)
            if datos is not None:
                resultados.append(datos);estructuras[anio]=estructura
        except Exception as error:
            texto=f'{anio}: {type(error).__name__}: {error}'
            informe.append('ERROR: '+texto);print(texto)
    informe.append('\nCAMBIOS ESTRUCTURALES ENTRE AÑOS:')
    for anterior,actual in zip(ANIOS,ANIOS[1:]):
        if anterior in estructuras and actual in estructuras:
            informe.append(f'{anterior}→{actual}: {estructuras[anterior]} → {estructuras[actual]}')
    exito=len(resultados)==len(ANIOS)
    if exito:
        datos=pd.concat(resultados,ignore_index=True).sort_values(COLUMNAS[:-1],kind='stable').reset_index(drop=True)
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
        informe.append(f'RESULTADO GLOBAL: NO VALIDADO; {len(resultados)}/16 años válidos; CSV no generado ni actualizado.')
        if SALIDA.exists():
            informe.append('Un archivo anterior no queda validado por esta ejecución.')
    REPORTE.parent.mkdir(parents=True,exist_ok=True)
    REPORTE.write_text('\n'.join(informe)+'\n',encoding='utf-8')
    print(informe[-1]);print(f'Reporte: {REPORTE}')
    if not exito:
        raise SystemExit(1)


if __name__=='__main__':
    main()
