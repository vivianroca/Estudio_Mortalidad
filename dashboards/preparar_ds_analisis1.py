"""Recalcula datos locales para ds_analisis1.qmd; no modifica los CSV originales."""
import hashlib
import json
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DASH = Path(__file__).resolve().parent
YEARS = list(range(2009,2025))

def norm(s):
    return ''.join(c for c in unicodedata.normalize('NFD',str(s)) if not unicodedata.combining(c)).strip().casefold()

def records(df):
    return json.loads(df.to_json(orient='records',force_ascii=False))

def main():
    paths={i:ROOT/f'data/processed/ind{i}_2009_2024.csv' for i in ('002','004','006')}
    paths['catalogo']=ROOT/'catalog/catalogo_cie10.csv'
    hashes={k:hashlib.sha256(p.read_bytes()).hexdigest() for k,p in paths.items()}
    d={i:pd.read_csv(paths[i],encoding='utf-8-sig') for i in ('002','004','006')}
    cat=pd.read_csv(paths['catalogo'],encoding='utf-8-sig',dtype={'codigo':str})
    assert cat.codigo.is_unique and cat.descripcion.is_unique
    official={norm(v):v for v in d['006'].depto_residencia.unique()}
    for i,t in d.items():
        assert set(t.anio)==set(YEARS) and not t.isna().any().any()
        assert t.conteo.ge(0).all() and (t.conteo%1==0).all()
        t['departamento']=t.depto_residencia.map(lambda v:official.get(norm(v)))
        assert t.departamento.notna().all()
    a2=d['002'].groupby(['anio','departamento']).conteo.sum().sort_index()
    a4=d['004'].groupby(['anio','departamento']).conteo.sum().sort_index()
    a6=d['006'].groupby(['anio','departamento']).conteo.sum().sort_index()
    pd.testing.assert_series_equal(a2,a4);pd.testing.assert_series_equal(a4,a6)
    b4=d['004'].groupby(['anio','departamento','sexo']).conteo.sum().sort_index()
    b6=d['006'].groupby(['anio','departamento','sexo']).conteo.sum().sort_index()
    pd.testing.assert_series_equal(b4,b6)
    catalog=dict(zip(cat.codigo,cat.descripcion))
    assert set(d['004'].codigo)<=set(catalog)
    res=d['006'].groupby(['anio','departamento','sexo'],as_index=False).conteo.sum()
    causes=d['004'].groupby(['anio','departamento','sexo','codigo'],as_index=False).conteo.sum()
    # Conservar solo combinaciones publicadas. No rellenar categorías ausentes con cero.
    ages=d['006'].assign(grupo_edad=d['006'].grupo_edad.str.strip()).groupby(
        ['anio','sexo','grupo_edad'],as_index=False).conteo.sum()
    mobility=d['002'].groupby(['anio','departamento','depto_ocurrencia'],as_index=False).conteo.sum()
    payload={
        'years':YEARS,'departments':sorted(official.values()),'labels':catalog,
        'residence':records(res),'causes':records(causes),'ages':records(ages),'mobility':records(mobility),
        'validation':{'annual_departments_compared':len(a2),'sex_margins_compared':len(b4),
                      'max_difference':0,'ind004_repeated_dimension_keys':int(d['004'].duplicated(
                          ['anio','departamento','sexo','codigo','causa_muerte','tiene_cie10']).sum())},
        'sources':[{'archivo':str(p.relative_to(ROOT)),'sha256':hashes[i],
                    'filas':len(d[i]) if i in d else len(cat)} for i,p in paths.items()],
        'generated':datetime.now().astimezone().isoformat(timespec='seconds'),
    }
    # Validación de conservación después de la agregación para visualizaciones.
    assert int(res.conteo.sum())==int(causes.conteo.sum())==int(mobility.conteo.sum())==1363143
    assert len(ages)==16*2*22
    for k,p in paths.items():
        assert hashlib.sha256(p.read_bytes()).hexdigest()==hashes[k]
    serialized=json.dumps(payload,ensure_ascii=False,separators=(',',':'),allow_nan=False).replace('</','<\\/')
    target=DASH/'ds_analisis1.qmd'
    template=target.read_text(encoding='utf-8')
    start='<!-- DATOS_INICIO -->';end='<!-- DATOS_FIN -->'
    assert template.count(start)==1 and template.count(end)==1
    before=template.split(start)[0];after=template.split(end)[1]
    target.write_text(before+start+'\n```{=html}\n<script id="datos-analisis" type="application/json">'+serialized+'</script>\n```\n'+end+after,encoding='utf-8')
    print('Dashboard actualizado:',target)
    print('Validado: 1.363.143 defunciones; 384 márgenes año/residencia y 768 márgenes año/residencia/sexo coinciden.')

if __name__=='__main__':
    main()
