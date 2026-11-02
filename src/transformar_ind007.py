"""IND-007, 2009–2024 sin 2011: prevalidar códigos y reproducir edad/causa/sexo sin imputar."""

import csv
import hashlib
import re
import unicodedata
from collections import defaultdict
from pathlib import Path

import openpyxl
import pandas as pd

from transformar_ind002 import entero, unico
from transformar_ind004 import leer_hoja

RAIZ = Path(__file__).resolve().parents[1]
# Decisión del proyecto: excluir solo IND-007 de 2011, sin corregir el original.
# En la fila 439 figura «83 - 84 años» dentro del bloque «80 - 84 años»:
# 321 defunciones (171 hombres, 150 mujeres). Los intervalos predominantes
# desde 5–9 hasta 80–84 abarcan cinco años. Posible errata no confirmada.
ANIOS = tuple(anio for anio in range(2009,2025) if anio != 2011)
SALIDA = RAIZ / 'data/processed/ind007_2009_2024_sin_2011.csv'
REPORTE = RAIZ / 'reports/estructura/validacion_ind007_2009_2024_sin_2011.txt'
COLUMNAS = ['anio','edad','causa_muerte','sexo','conteo','tiene_cie10','codigo']
SEXOS = ('Hombres','Mujeres')


def clave(texto):
    return ' '.join(unicodedata.normalize('NFC',str(texto)).split()).casefold()


def codigos_finales():
    final=defaultdict(set)
    with (RAIZ/'catalog/catalogo_cie10.csv').open(encoding='utf-8-sig',newline='') as f:
        for r in csv.DictReader(f):
            final[clave(r['descripcion'])].add(r['codigo'])
    if any(len(c)!=1 for c in final.values()):
        raise ValueError('Catálogo final ambiguo')
    lookup={k:set(v) for k,v in final.items()}
    fuentes=[('reports/estructura/propuesta_cat_cie10.xlsx','rev_final','DESCRIPCION','DESCRIPCION_FINAL'),
             ('catalog/posible_cat.xlsx','Revision_IND007','descripcion_original','descripcion_final')]
    for archivo,hoja,original,oficial in fuentes:
        b=openpyxl.load_workbook(RAIZ/archivo,read_only=True,data_only=False)
        try:
            it=b[hoja].iter_rows(values_only=True);h=next(it)
            for r in it:
                if not any(v is not None for v in r):
                    continue
                codes=final.get(clave(r[h.index(oficial)]),set())
                if len(codes)!=1:
                    raise ValueError(f'{archivo}: descripción oficial sin código único')
                lookup.setdefault(clave(r[h.index(original)]),set()).update(codes)
        finally:
            b.close()
    return lookup


def cargar(anio,nombre):
    ruta,huella,matriz,dimensiones=leer_hoja(anio,nombre)
    r_enc=unico([r for r,fila in enumerate(matriz) if all(t in fila for t in
                ('Edad','Causa de muerte','Total',*SEXOS))],'Encabezado IND-007')
    cols={t:unico([i for i,v in enumerate(matriz[r_enc]) if v==t],t)
          for t in ('Edad','Causa de muerte','Total',*SEXOS)}
    cie=[i for i,v in enumerate(matriz[r_enc]) if v=='Código CIE-10']
    filas=[]
    for r in range(r_enc+1,len(matriz)):
        fila=matriz[r]
        if not any(v!='' for v in fila):
            continue
        edad,causa=fila[cols['Edad']],fila[cols['Causa de muerte']]
        if isinstance(edad,str) and clave(edad).startswith(('fuente:','nota:','notas:')):
            continue
        if not isinstance(edad,str) or not edad.strip() or not isinstance(causa,str) or not causa.strip():
            raise ValueError(f'{anio}, fila {r+1}: edad/causa vacía o no textual')
        valores={s:entero(fila[cols[s]],f'{anio}, fila {r+1}, {s}') for s in ('Total',*SEXOS)}
        code=str(fila[cie[0]]).strip() if cie else ''
        flag=int(bool(re.fullmatch(r'[A-Z]\d{2}(?:\.\d+)?(?:\s*[-–]\s*[A-Z]\d{2}(?:\.\d+)?)?',code)))
        if any(v!='' for c,v in enumerate(fila) if c not in set(cols.values())|set(cie)):
            raise ValueError(f'{anio}, fila {r+1}: celdas fuera del cuadro')
        filas.append(dict(fila=r+1,edad=edad,causa=causa,v=valores,flag=flag))
    return ruta,huella,dimensiones,r_enc,filas


def transformar(anio,nombre,libro,lookup,informe):
    ruta,huella,dimensiones,encabezado,filas=libro
    errores=[]
    informe.extend([f'\nAÑO: {anio}; hoja: {nombre}; archivo: {ruta.name}',
                    f'Dimensiones: {dimensiones}; encabezado: fila {encabezado+1}; SHA-256: {huella}.'])
    por_edad=defaultdict(list)
    for f in filas:
        por_edad[f['edad']].append(f)
        if f['v']['Hombres']+f['v']['Mujeres']!=f['v']['Total']:
            errores.append(f'Fila {f["fila"]}: suma de sexos distinta del Total')
    totales={}
    for edad,grupo in por_edad.items():
        ts=[f for f in grupo if f['causa']=='Todas las causas']
        if len(ts)!=1:
            errores.append(f'{edad!r}: {len(ts)} totales Todas las causas; filas={[f["fila"] for f in grupo]}. No se reasignan filas a otro grupo.')
            continue
        totales[edad]=ts[0]['v']
        for sexo in ('Total',*SEXOS):
            suma=sum(f['v'][sexo] for f in grupo if f['causa']!='Todas las causas')
            total=ts[0]['v'][sexo]
            informe.append(f'{edad!r}/{sexo}: total original={total}; suma de causas={suma}; diferencia={suma-total}.')
            if suma!=total:
                errores.append(f'{edad!r}/{sexo}: causas={suma}, total={total}, diferencia={suma-total}.')
    analiticas=[f for f in filas if f['edad']!='Todas las edades' and f['causa']!='Todas las causas']
    ancho=pd.DataFrame([dict(_fila=f['fila'],edad=f['edad'],causa_muerte=f['causa'],
                            tiene_cie10=f['flag'],codigo=next(iter(lookup[clave(f['causa'])])),
                            **{s:f['v'][s] for s in SEXOS}) for f in analiticas])
    largo=ancho.melt(id_vars=['_fila','edad','causa_muerte','tiene_cie10','codigo'],
                      value_vars=list(SEXOS),var_name='sexo',value_name='conteo')
    largo.insert(0,'anio',anio)
    fuente={(f['fila'],s):(f['edad'],f['causa'],f['v'][s]) for f in analiticas for s in SEXOS}
    reconstruido={(r._0,r.sexo):(r.edad,r.causa_muerte,r.conteo)
                  for r in largo[['_fila','sexo','edad','causa_muerte','conteo']].itertuples(index=False)}
    if fuente!=reconstruido or len(largo)!=len(fuente):
        errores.append('Intersecciones incorrectas o filas perdidas después de melt')
    if 'Todas las edades' not in totales:
        errores.append('Falta total nacional Todas las edades')
    else:
        for sexo in ('Total',*SEXOS):
            calculado=int(largo.conteo.sum() if sexo=='Total' else largo.loc[largo.sexo==sexo,'conteo'].sum())
            publicado=totales['Todas las edades'][sexo]
            informe.append(f'TOTAL GENERAL/{sexo}: publicado={publicado}; transformado={calculado}; diferencia={calculado-publicado}.')
            if calculado!=publicado:
                errores.append(f'Total general no coincide: {sexo}')
            margen=sum(v[sexo] for e,v in totales.items() if e!='Todas las edades')
            if margen!=publicado:
                errores.append(f'Márgenes de edad no suman el total nacional: {sexo}, diferencia={margen-publicado}')
    if hashlib.sha256(ruta.read_bytes()).hexdigest()!=huella:
        errores.append('Hash original modificado')
    informe.append(f'Intersecciones comprobadas: {len(largo)}; códigos finales vacíos: {largo.codigo.isna().sum()}.')
    informe.extend('ERROR: '+e for e in errores)
    informe.append(f'RESULTADO {anio}: '+('NO VALIDADO.' if errores else 'VALIDADO.'))
    print(f'{anio}: {len(largo)} registros; '+('NO VALIDADO' if errores else 'VALIDADO'))
    for error in errores:
        print('  '+error)
    return None if errores else largo[COLUMNAS]


def main():
    lookup=codigos_finales()
    with (RAIZ/'reports/estructura/catalogo_indicadores.csv').open(encoding='utf-8-sig') as f:
        indicador=unico([r for r in csv.DictReader(f) if r['id_provisional']=='IND-007'],'IND-007')
    libros={};errores=[];faltantes=defaultdict(int)
    for anio in ANIOS:
        try:
            libros[anio]=cargar(anio,indicador[str(anio)])
            for f in libros[anio][-1]:
                if f['causa']=='Todas las causas':
                    continue
                if len(lookup.get(clave(f['causa']),set()))!=1:
                    faltantes.setdefault(f['causa'],0)
                    if f['edad']!='Todas las edades':
                        faltantes[f['causa']]+=f['v']['Total']
        except Exception as error:
            errores.append(f'{anio}: {type(error).__name__}: {error}')
    if faltantes:
        print('PROCESO DETENIDO: descripciones sin código final único; sin exportar archivos.')
        for causa,total in sorted(faltantes.items(),key=lambda p:(-p[1],p[0])):
            print(f'{causa} | {total}')
        raise SystemExit(1)
    print('Prevalidación: todas las descripciones tienen código final único.')
    informe=['VALIDACIÓN IND-007 — 2009–2024, EXCLUYENDO 2011',
             'Decisión expresa del proyecto: excluir 2011 únicamente para IND-007; no se abre ese libro en esta ejecución.',
             'Antecedente: fila 439 de 2011, Insuficiencia cardíaca, etiqueta 83 - 84 años dentro del bloque 80 - 84 años; 321 defunciones (171 hombres y 150 mujeres).',
             'Los intervalos predominantes desde 5–9 hasta 80–84 son quinquenales. La etiqueta discrepante es una posible errata; no se confirmó ni corrigió.',
             'La exclusión de 2011 no significa cero defunciones ni elimina el año de otros indicadores.',
             'Código final mediante catálogo único y variantes aprobadas; sin inferencias por similitud.',
             'Conservar etiquetas originales de edad: hay días, intervalos, edades simples e Ignorado. No se convierten intervalos a una edad puntual.',
             'No se rechazan descripciones iguales con códigos originales diferentes ni se fusionan sus filas.',
             'tiene_cie10 registra código explícito en la fila original; codigo contiene la asignación del catálogo final.',
             'Todas las edades, Todas las causas y Total se excluyen del CSV y se usan para comprobar sumas.',
             'No se completan causas no listadas con cero; Otras causas es residual variable.',
             'Guiones, vacíos y no numéricos bloquean la salida. No se corrigen etiquetas originales automáticamente.']
    informe.extend('ERROR: '+e for e in errores)
    resultados=[]
    for anio,libro in libros.items():
        try:
            datos=transformar(anio,indicador[str(anio)],libro,lookup,informe)
            if datos is not None:
                resultados.append(datos)
        except Exception as error:
            informe.append(f'ERROR {anio}: {type(error).__name__}: {error}')
    exito=len(resultados)==len(ANIOS)
    if exito:
        datos=pd.concat(resultados,ignore_index=True).sort_values(['anio','edad','causa_muerte','sexo'],kind='stable').reset_index(drop=True)
        temporal=SALIDA.with_suffix('.tmp.csv')
        try:
            SALIDA.parent.mkdir(parents=True,exist_ok=True)
            datos.to_csv(temporal,index=False,encoding='utf-8-sig')
            pd.testing.assert_frame_equal(datos,pd.read_csv(temporal,encoding='utf-8-sig',dtype={'codigo':str}))
            temporal.replace(SALIDA)
            informe.append(f'RESULTADO GLOBAL: VALIDADO; {len(datos)} registros; {SALIDA}')
        finally:
            temporal.unlink(missing_ok=True)
    else:
        informe.append(f'RESULTADO GLOBAL: NO VALIDADO; {len(resultados)}/{len(ANIOS)} años válidos. No se generó ni actualizó el CSV.')
        if SALIDA.exists():
            informe.append('Existe un archivo anterior; no está validado por esta ejecución.')
    REPORTE.parent.mkdir(parents=True,exist_ok=True)
    REPORTE.write_text('\n'.join(informe)+'\n',encoding='utf-8')
    print(informe[-1]);print(f'Diagnóstico: {REPORTE}')
    if not exito:
        raise SystemExit(1)


if __name__=='__main__':
    main()
