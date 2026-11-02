"""Agregados y figuras del EDA Quarto. Lectura única; no ETL ni escritura de datos."""
from hashlib import sha256
import json
from pathlib import Path
import textwrap

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.offline import get_plotlyjs
from IPython.display import display, HTML, Markdown

ROOT = Path(__file__).resolve().parents[1]
AGES = ['<1 año', '1–4', '5–14', '15–24', '25–44', '45–64', '65–79', '80+', 'Ignorada']
BROAD = ['<1 año', '1–14', '15–44', '45–64', '65+', 'Ignorada']
PERIODS = ['Prepandemia (2009–2019)', 'Pandemia (2020–2021)', 'Pospandemia (2022–2024)']
CV = 'Enfermedades Cardiovasculares (ENT)'
ILL = 'Signos y Causas Mal Definidas'
SOURCE = 'Fuente: microdatos INE, merge depurado 2009–2024. Año de registro. Conteos o participaciones; no tasas.'


def number(value):
    return f'{int(value):,}'.replace(',', '.')


def percent(value):
    return f'{float(value):.2f}'.replace('.', ',') + ' %'


def wrap(value, width=48):
    return '<br>'.join(textwrap.wrap(str(value), width))


class EDA:
    def __init__(self):
        self.path = ROOT / 'data/processed/merge_final.parquet'
        self.protected = {p: sha256(p.read_bytes()).hexdigest() for p in
                          [self.path, ROOT / 'docu/diccionario_datos.md', ROOT / 'docu/diccionario_merge_final.xlsx']}
        cols = ['anio', 'sexo', 'edad', 'periodo_edad', 'depto_res', 'causa',
                'enfermedad_normalizada', 'grupo_epidemiologico']
        d = pd.read_parquet(self.path, columns=cols)  # Única carga del Parquet.
        self.n = len(d)
        assert self.n == 1323982, 'El total no coincide con la auditoría de referencia.'
        assert set(d.anio.unique()) == set(range(2009, 2025))
        assert d.grupo_epidemiologico.nunique() == 19
        assert not d[['anio', 'enfermedad_normalizada', 'grupo_epidemiologico']].isna().any().any()
        d['sexo_nombre'] = d.sexo.map({1: 'Hombres', 2: 'Mujeres', 9: 'Ignorado'}).fillna('Otro/no identificado')
        ignored = d.periodo_edad.eq(9) | d.edad.eq(999) | d.edad.isna()
        infant = d.periodo_edad.isin([1, 2]) & ~ignored
        # Los datos auditados no contienen 12 meses; no interpretar nuevos casos así sin revisar.
        assert not (d.periodo_edad.eq(2) & d.edad.ge(12) & ~ignored).any()
        years = d.edad.where(d.periodo_edad.eq(3) & ~ignored)
        detailed = pd.cut(years, [0, 4, 14, 24, 44, 64, 79, np.inf], labels=AGES[1:-1])
        d['grupo_edad'] = detailed.astype('string').mask(infant, AGES[0]).mask(ignored, AGES[-1])
        broad = pd.cut(years, [0, 14, 44, 64, np.inf], labels=BROAD[1:-1])
        d['edad_amplia'] = broad.astype('string').mask(infant, BROAD[0]).mask(ignored, BROAD[-1])
        assert not d.grupo_edad.isna().any() and not d.edad_amplia.isna().any()
        d['periodo'] = np.select([d.anio.le(2019), d.anio.le(2021)], PERIODS[:2], default=PERIODS[2])
        self.annual = d.groupby('anio').size().rename('registros').reset_index()
        self.annual['variacion_pct'] = self.annual.registros.pct_change() * 100
        self.annual['cambio_absoluto'] = self.annual.registros.diff()
        annual_den = self.annual.set_index('anio').registros
        assert int(self.annual.registros.sum()) == self.n

        def aggregate(keys):
            return d.groupby(keys, observed=True).size().rename('registros').reset_index()
        self.sex = aggregate(['sexo_nombre']); self.sex['porcentaje'] = self.sex.registros / self.n * 100
        self.age = aggregate(['grupo_edad']); self.age['porcentaje'] = self.age.registros / self.n * 100
        self.pyramid = aggregate(['grupo_edad', 'sexo_nombre'])
        self.sex_year = aggregate(['anio', 'sexo_nombre'])
        self.age_year = aggregate(['anio', 'grupo_edad'])
        for frame in [self.sex_year, self.age_year]:
            frame['porcentaje'] = frame.registros / frame.anio.map(annual_den) * 100
        self.groups = aggregate(['grupo_epidemiologico']).sort_values('registros', ascending=False)
        self.groups['porcentaje'] = self.groups.registros / self.n * 100
        self.top_groups = self.groups.grupo_epidemiologico.head(10).tolist()
        self.group_year = aggregate(['anio', 'periodo', 'grupo_epidemiologico'])
        self.group_year['porcentaje'] = self.group_year.registros / self.group_year.anio.map(annual_den) * 100
        self.diseases = aggregate(['enfermedad_normalizada']).sort_values(['registros', 'enfermedad_normalizada'], ascending=[False, True])
        self.diseases['porcentaje'] = self.diseases.registros / self.n * 100
        self.top_diseases = self.diseases.enfermedad_normalizada.head(20).tolist()
        self.disease_year = aggregate(['anio', 'periodo', 'enfermedad_normalizada'])
        self.disease_year['porcentaje'] = self.disease_year.registros / self.disease_year.anio.map(annual_den) * 100
        self.disease_sex = aggregate(['enfermedad_normalizada', 'sexo_nombre'])
        sex_den = self.sex.set_index('sexo_nombre').registros
        self.disease_sex['porcentaje'] = self.disease_sex.registros / self.disease_sex.sexo_nombre.map(sex_den) * 100
        self.disease_age = aggregate(['enfermedad_normalizada', 'edad_amplia'])
        age_den = d.edad_amplia.value_counts()
        self.disease_age['porcentaje'] = self.disease_age.registros / self.disease_age.edad_amplia.map(age_den) * 100
        self.age_leaders = self.disease_age.sort_values('registros', ascending=False).groupby('edad_amplia', sort=False).head(3)
        self.age_candidates = self.disease_age.sort_values('registros', ascending=False).groupby('edad_amplia', sort=False).head(5).enfermedad_normalizada.drop_duplicates().tolist()
        assert int(self.groups.registros.sum()) == self.n
        assert int(self.diseases.registros.sum()) == self.n
        assert int(self.sex.registros.sum()) == int(self.age.registros.sum()) == self.n

        annual_period = self.annual.copy()
        annual_period['periodo'] = np.select([annual_period.anio.le(2019), annual_period.anio.le(2021)], PERIODS[:2], default=PERIODS[2])
        self.periods = annual_period.groupby('periodo').agg(anios=('anio', 'size'), registros=('registros', 'sum'), promedio_anual=('registros', 'mean')).reindex(PERIODS).reset_index()
        assert int(self.periods.registros.sum()) == self.n

        # Cuadrículas completas: ceros solo en categorías con denominador observado.
        def period_composition(frame, label):
            grid = pd.MultiIndex.from_product([range(2009, 2025), frame[label].unique()], names=['anio', label])
            complete = frame.set_index(['anio', label]).porcentaje.reindex(grid, fill_value=0).reset_index()
            complete['periodo'] = np.select([complete.anio.le(2019), complete.anio.le(2021)], PERIODS[:2], default=PERIODS[2])
            return complete.groupby(['periodo', label]).porcentaje.mean().reset_index()
        self.group_period = period_composition(self.group_year, 'grupo_epidemiologico')
        self.disease_period = period_composition(self.disease_year, 'enfermedad_normalizada')
        assert np.allclose(self.group_year.groupby('anio').porcentaje.sum(), 100)
        assert np.allclose(self.group_period.groupby('periodo').porcentaje.sum(), 100)
        indicators = {
            ILL: d.grupo_epidemiologico.eq(ILL), 'Código R99X': d.causa.eq('R99X'),
            'Edad ignorada': ignored, 'Residencia ignorada': d.depto_res.eq(99),
        }
        self.quality = pd.concat([
            pd.DataFrame({'anio': d.anio, 'indicador': name, 'marca': mask}).groupby(['anio', 'indicador']).marca.sum().rename('registros').reset_index()
            for name, mask in indicators.items()], ignore_index=True)
        self.quality['porcentaje'] = self.quality.registros / self.quality.anio.map(annual_den) * 100
        self.extreme = int((d.periodo_edad.eq(3) & d.edad.gt(110) & d.edad.ne(999)).sum())
        self.ignored = int(ignored.sum())
        self.nonspecific = int(d.causa.eq('R99X').sum())

        geo_path = ROOT / 'data/geo/departamentos_guatemala.geojson'
        self.geo = json.loads(geo_path.read_text())
        self.protected[geo_path] = sha256(geo_path.read_bytes()).hexdigest()
        features = self.geo['features']
        assert len(features) == 22 and {f['properties']['depto_res'] for f in features} == set(range(1, 23))
        self.names = {f['properties']['depto_res']: f['properties']['Departamento'] for f in features}
        correspondence = pd.read_csv(ROOT / 'data/geo/correspondencia_departamentos.csv')
        assert correspondence.depto_res.is_unique and set(correspondence.depto_res) == set(range(1, 23))
        for row in correspondence.itertuples():
            assert row.id_original == row.depto_res * 100
            assert self.names[row.depto_res] == row.departamento_fuente
            assert row.validado_codigo and row.validado_nombre
        domestic = d.depto_res.between(1, 22)
        self.excluded = d.loc[~domestic].groupby('depto_res').size().rename('registros').reset_index()
        self.excluded['residencia'] = self.excluded.depto_res.map({23: 'Extranjero', 99: 'Ignorado'})
        self.excluded['porcentaje_total'] = self.excluded.registros / self.n * 100
        self.territory = d.loc[domestic].groupby('depto_res').size().reindex(range(1, 23), fill_value=0).rename('registros').reset_index()
        self.territory['departamento'] = self.territory.depto_res.map(self.names)
        self.dept_groups = d.loc[domestic].groupby(['depto_res', 'grupo_epidemiologico']).size().rename('conteo_grupo').reset_index()
        den = self.territory.set_index('depto_res').registros
        self.dept_groups['porcentaje'] = self.dept_groups.conteo_grupo / self.dept_groups.depto_res.map(den) * 100
        self.dept_groups['departamento'] = self.dept_groups.depto_res.map(self.names)
        for name, column in [(CV, 'cardiovascular_pct'), (ILL, 'mal_definidas_pct')]:
            values = self.dept_groups.loc[self.dept_groups.grupo_epidemiologico.eq(name)].set_index('depto_res')
            self.territory[column] = self.territory.depto_res.map(values.porcentaje).fillna(0)
            self.territory[column.replace('_pct', '_n')] = self.territory.depto_res.map(values.conteo_grupo).fillna(0).astype(int)
        assert int(self.territory.registros.sum()) + int(self.excluded.registros.sum()) == self.n
        assert np.allclose(self.dept_groups.groupby('depto_res').porcentaje.sum(), 100)
        self.figures = 0
        self.d = d

    def table(self, frame, rename=None, max_rows=None):
        frame = frame.rename(columns=rename or {}).copy()
        if max_rows is not None:
            frame = frame.head(max_rows)
        display(HTML(frame.to_html(index=False, border=0, classes=['table', 'table-striped', 'table-sm'],
                                  na_rep='No aplica', float_format=lambda x: f'{x:,.2f}')))

    def javascript(self):
        display(HTML('<script>' + get_plotlyjs() + '</script>'))

    def emit(self, fig, caption):
        self.figures += 1
        fig.update_layout(template='plotly_white', font={'family': 'Arial', 'size': 13},
                          margin={'t': 80, 'b': 65, 'l': 60, 'r': 25})
        post_script = "window.edaReady = window.edaReady || []; window.edaReady.push('{plot_id}');"
        if fig.data[0].type == 'choroplethmap':
            west,south,east,north = self.geo['bbox']
            bounds = json.dumps([[west-.1,south-.1],[east+.1,north+.1]])
            post_script += """
const el = document.getElementById('{plot_id}');
const map = el._fullLayout.map._subplot.map;
const fit = () => { map.resize(); map.fitBounds(BOUNDS, {padding: 18, duration: 0}); };
requestAnimationFrame(fit);
el._edaResizeObserver = new ResizeObserver(() => requestAnimationFrame(fit));
el._edaResizeObserver.observe(el);
""".replace('BOUNDS', bounds)
        display(HTML(fig.to_html(full_html=False, include_plotlyjs=False,
                                div_id=f'eda-fig-{self.figures}',
                                config={'responsive': True, 'displaylogo': False},
                                post_script=post_script)))
        display(Markdown(f'*{caption} {SOURCE}*'))

    def bar(self, frame, label, title, percent_col=False, height=650):
        frame = frame.sort_values('registros', ascending=True).copy()
        frame['rotulo'] = frame[label].map(wrap)
        x = 'porcentaje' if percent_col else 'registros'
        fig = px.bar(frame, x=x, y='rotulo', orientation='h', custom_data=[label, 'registros'], title=title,
                     labels={'rotulo': '', x: 'Participación (%)' if percent_col else 'Registros'}, height=height)
        fig.update_traces(hovertemplate='%{customdata[0]}<br>Registros: %{customdata[1]:,}<br>Valor: %{x:,.2f}<extra></extra>')
        return fig

    def heat(self, frame, row, col, value, rows, cols, title, height=650, percentage=True):
        matrix = frame.pivot_table(index=row, columns=col, values=value, aggfunc='sum', fill_value=0).reindex(index=rows, columns=cols, fill_value=0)
        z = matrix.to_numpy(dtype=float)
        fig = go.Figure(go.Heatmap(z=z, x=[wrap(v, 26) for v in cols], y=[wrap(v, 52) for v in rows],
                                  colorscale='Blues', zmin=0, colorbar={'title': '%' if percentage else 'Registros'},
                                  hovertemplate='%{y}<br>%{x}<br>%{z:,.2f}' + (' %' if percentage else ' registros') + '<extra></extra>'))
        fig.update_layout(title=title, height=height, yaxis={'autorange': 'reversed', 'automargin': True})
        return fig

    def show(self, name):
        if name == 'annual':
            fig = px.line(self.annual, x='anio', y='registros', markers=True, title='Defunciones registradas por año', labels={'anio': 'Año de registro', 'registros': 'Registros'})
            self.emit(fig, 'Total anual dentro de la cohorte cargada.')
        elif name == 'yoy':
            self.emit(px.bar(self.annual.dropna(), x='anio', y='variacion_pct', title='Cambio interanual del conteo registrado', labels={'anio': 'Año de registro', 'variacion_pct': 'Variación respecto al año anterior (%)'}), 'Denominador: conteo del año anterior. En 2009 no se dispone de 2008.')
        elif name == 'periods':
            self.emit(px.bar(self.periods, x='periodo', y='promedio_anual', title='Comparación de períodos: promedio anual', labels={'periodo': '', 'promedio_anual': 'Registros por año (media)'}), 'Sumas divididas entre los años de cada período; evita comparar directamente intervalos de 11, 2 y 3 años.')
        elif name == 'sex':
            self.emit(self.bar(self.sex, 'sexo_nombre', 'Composición de las defunciones registradas por sexo', True, 400), 'Denominador: todos los registros; el porcentaje no estima riesgo por sexo.')
        elif name == 'age':
            self.emit(self.bar(self.age, 'grupo_edad', 'Distribución por grupos de edad', True), 'Menores de un año reúne días y meses; edad ignorada permanece explícita.')
        elif name == 'pyramid':
            fig = go.Figure()
            for name_, sign in [('Hombres', -1), ('Mujeres', 1)]:
                v = self.pyramid.loc[self.pyramid.sexo_nombre.eq(name_)].set_index('grupo_edad').registros.reindex(AGES, fill_value=0)
                fig.add_bar(name=name_, y=AGES, x=v.to_numpy() * sign, orientation='h', customdata=v.to_numpy(), hovertemplate=name_+'<br>%{y}<br>%{customdata:,} registros<extra></extra>')
            maximum = int(self.pyramid.registros.max())
            ticks = np.linspace(-maximum, maximum, 7)
            fig.update_layout(title='Pirámide de defunciones por sexo y edad', barmode='relative', height=600,
                              xaxis={'title': 'Registros (valor absoluto a ambos lados)', 'tickvals': ticks, 'ticktext': [number(abs(v)) for v in ticks]}, yaxis={'categoryorder': 'array', 'categoryarray': AGES})
            self.emit(fig, 'Frecuencias, no pirámide de población. Edad ignorada se muestra como categoría; el signo negativo solo permite dibujar hombres a la izquierda.')
        elif name in ['sex_year', 'age_year']:
            frame, label = (self.sex_year, 'sexo_nombre') if name == 'sex_year' else (self.age_year, 'grupo_edad')
            fig = px.bar(frame, x='anio', y='porcentaje', color=label, barmode='stack', category_orders={'grupo_edad': AGES},
                         title='Composición anual por '+('sexo' if name=='sex_year' else 'edad'), labels={'anio':'Año de registro','porcentaje':'Porcentaje anual',label:''})
            self.emit(fig, 'Denominador: total de registros de cada año, incluidas categorías ignoradas.')
        elif name == 'groups':
            self.emit(self.bar(self.groups, 'grupo_epidemiologico', 'Los 19 grupos epidemiológicos originales', True, 1000), 'Denominador: total acumulado; nombres de la v2 conservados sin fusionar.')
        elif name == 'group_year':
            frame = self.group_year.loc[self.group_year.grupo_epidemiologico.isin(self.top_groups)].copy()
            frame['grupo'] = frame.grupo_epidemiologico.map(lambda v: wrap(v, 40))
            fig = px.line(frame, x='anio', y='registros', color='grupo', markers=True, title='Trayectoria de los diez grupos principales', labels={'anio':'Año de registro','registros':'Registros','grupo':''}, height=780)
            fig.update_layout(legend={'orientation':'h','y':-0.22,'x':0})
            self.emit(fig, 'Los diez grupos se seleccionaron por frecuencia acumulada global, manteniendo la misma selección en todos los años. Use la leyenda para aislar series.')
        elif name in ['group_share', 'group_period']:
            if name == 'group_share':
                fig=self.heat(self.group_year,'grupo_epidemiologico','anio','porcentaje',self.groups.grupo_epidemiologico.tolist(),list(range(2009,2025)),'Composición anual: los 19 grupos',1050)
            else:
                fig=self.heat(self.group_period,'grupo_epidemiologico','periodo','porcentaje',self.groups.grupo_epidemiologico.tolist(),PERIODS,'Media de las participaciones anuales por período',1050)
            self.emit(fig, 'Denominador de cada participación: total anual de registros. En períodos se promedian las participaciones anuales con igual peso por año; cada columna suma 100 %.')
        elif name == 'diseases':
            self.emit(self.bar(self.diseases.head(20),'enfermedad_normalizada','Veinte enfermedades normalizadas más frecuentes',False,1150),'Los rótulos se conservan exactamente como están en el merge; incluyen categorías generales de precisión desigual.')
        elif name == 'disease_year':
            self.emit(self.heat(self.disease_year,'enfermedad_normalizada','anio','registros',self.top_diseases[:10],list(range(2009,2025)),'Evolución anual: diez enfermedades principales',750,False),'Selección fija según frecuencia acumulada global. Cada celda cuenta registros; cero indica que esa etiqueta no aparece en la base filtrada.')
        elif name == 'disease_sex':
            frame=self.disease_sex.loc[self.disease_sex.enfermedad_normalizada.isin(self.top_diseases[:10]) & self.disease_sex.sexo_nombre.isin(['Hombres','Mujeres'])].copy()
            frame['enfermedad']=frame.enfermedad_normalizada.map(wrap)
            fig=px.bar(frame,x='porcentaje',y='enfermedad',color='sexo_nombre',orientation='h',barmode='group',custom_data=['registros'],height=850,title='Principales enfermedades: composición dentro de cada sexo',labels={'porcentaje':'% del total de registros de ese sexo','enfermedad':'','sexo_nombre':''})
            fig.update_traces(hovertemplate='%{y}<br>%{x:.2f} %<br>%{customdata[0]:,} registros<extra></extra>')
            self.emit(fig,'Denominador: total del sexo correspondiente, no el total de cada enfermedad. La selección es el top diez global.')
        elif name == 'disease_age':
            self.emit(self.heat(self.disease_age,'enfermedad_normalizada','edad_amplia','porcentaje',self.age_candidates,BROAD,'Enfermedades por grandes grupos etarios',max(800,len(self.age_candidates)*45)), 'Selección: unión de las cinco enfermedades más frecuentes dentro de cada grupo etario. Denominador: todos los registros de ese grupo, incluida edad ignorada en su propia columna.')
        elif name == 'disease_period':
            self.emit(self.heat(self.disease_period,'enfermedad_normalizada','periodo','porcentaje',self.top_diseases[:15],PERIODS,'Principales enfermedades: composición media anual por período',1050),'Denominador: total anual de registros, antes de promediar los porcentajes por período. Se muestran quince etiquetas; las columnas mostradas no suman 100 %.')
        elif name == 'quality':
            fig=px.line(self.quality,x='anio',y='porcentaje',color='indicador',markers=True,title='Cobertura y especificidad del registro: indicadores anuales',labels={'anio':'Año de registro','porcentaje':'% del total anual','indicador':''},height=650)
            fig.update_layout(legend={'orientation':'h','y':-0.22})
            self.emit(fig,'Cada indicador usa como denominador el total de su año. R99X pertenece al conjunto de causas mal definidas: las series se superponen y no deben sumarse.')
        elif name == 'territory_heat':
            self.emit(self.heat(self.dept_groups,'departamento','grupo_epidemiologico','porcentaje',[self.names[i] for i in range(1,23)],self.top_groups,'Composición territorial de los diez grupos principales',1000),'Denominador: todos los registros del departamento de residencia correspondiente. Los diez grupos son los principales a escala global; las filas mostradas pueden sumar menos de 100 %.')
        elif name.startswith('map_'):
            variable,title,label = {
                'map_count':('registros','Defunciones registradas por departamento','Registros acumulados'),
                'map_cv':('cardiovascular_pct','Participación cardiovascular dentro de cada departamento','Porcentaje dentro del departamento'),
                'map_ill':('mal_definidas_pct','Participación de causas mal definidas por departamento','Porcentaje dentro del departamento'),
            }[name]
            fig=px.choropleth_map(self.territory,geojson=self.geo,locations='depto_res',featureidkey='properties.depto_res',color=variable,
                                  hover_name='departamento',hover_data={'depto_res':True,'registros':':,',variable:':.2f' if variable!='registros' else ':,',
                                  'cardiovascular_n':True if name=='map_cv' else False,'mal_definidas_n':True if name=='map_ill' else False},
                                  map_style='white-bg',center={'lat':15.7,'lon':-90.3},zoom=5.6,height=700,
                                  color_continuous_scale='Blues',range_color=(0,float(self.territory[variable].max())),title=title,
                                  labels={variable:label,'registros':'Registros del departamento','depto_res':'Código residencia','cardiovascular_n':'Registros cardiovasculares','mal_definidas_n':'Registros mal definidos'})
            fig.update_layout(coloraxis_colorbar={'title':{'text':'Registros' if name=='map_count' else '% depto.'},'thickness':18,'len':.8})
            self.emit(fig,'Solo códigos 1–22. Fondo vacío y polígonos locales, sin teselas ni servidor Python. El indicador porcentual usa como denominador los registros del mismo departamento. No son tasas.')
        else:
            raise KeyError(name)

    def note(self, section):
        peak=self.annual.loc[self.annual.registros.idxmax()]
        if section=='intro':
            text=f'Se analizan **{number(self.n)} registros**, **{self.d.enfermedad_normalizada.nunique():,} enfermedades normalizadas**, **{len(self.groups)} grupos** y **{len(self.annual)} años de registro**, de {int(self.annual.anio.min())} a {int(self.annual.anio.max())}. Los conteos anuales y por grupo concilian con el total auditado.'
        elif section=='temporal':
            change=self.annual.loc[self.annual.variacion_pct.idxmax()]
            text=f'El mayor conteo se observa en **{int(peak.anio)} ({number(peak.registros)})**. El mayor incremento interanual corresponde a **{int(change.anio)} ({percent(change.variacion_pct)})**. Estas son diferencias observadas en la cohorte registrada: no estiman exceso de mortalidad ni atribuyen el cambio a una causa.'
        elif section=='demography':
            older=int(self.age.loc[self.age.grupo_edad.isin(['65–79','80+']),'registros'].sum())
            text=f'Las edades de 65 años o más suman **{number(older)} registros ({percent(older/self.n*100)})**. Hay **{number(self.ignored)} edades ignoradas** y **{number(self.extreme)} registros con más de 110 años** en período 3; se conservan, sin corregirlos ni tratarlos como errores probados. La pirámide caracteriza defunciones, no la población expuesta.'
        elif section=='groups':
            row=self.groups.iloc[0]
            text=f'El grupo con mayor frecuencia acumulada es **{row.grupo_epidemiologico}**, con **{number(row.registros)} registros ({percent(row.porcentaje)})**. La composición anual y las medias por período permiten observar cambios relativos, pero dependen de edad, cobertura y codificación; una participación menor no implica una reducción del riesgo poblacional.'
        elif section=='diseases':
            row=self.diseases.iloc[0]
            text=f'La etiqueta más frecuente es **{row.enfermedad_normalizada} ({number(row.registros)})**. El top veinte reúne **{percent(self.diseases.head(20).registros.sum()/self.n*100)}** del total. Categorías como R99X y otras etiquetas generales no tienen la misma precisión que un diagnóstico específico; no se fusionaron nombres similares ni se revalidaron las equivalencias clínicas.'
        elif section=='quality':
            text=f'R99X suma **{number(self.nonspecific)} registros ({percent(self.nonspecific/self.n*100)})**, incluyendo las reasignaciones previas documentadas. La edad y residencia ignoradas se mantienen en el denominador general. Variaciones de cobertura o especificidad pueden cambiar las composiciones sin reflejar cambios equivalentes en los fenómenos clínicos.'
        elif section=='territory':
            excluded=int(self.excluded.registros.sum())
            text=f'Los mapas incluyen **{number(self.territory.registros.sum())} registros** de los 22 departamentos. Se excluyen solo de las visualizaciones geográficas **{number(excluded)} ({percent(excluded/self.n*100)})** con residencia extranjera o ignorada, que permanecen en los análisis generales. Los conteos territoriales dependen del tamaño y composición poblacional; las proporciones describen únicamente las defunciones registradas en cada departamento.'
        else: raise KeyError(section)
        display(Markdown(text))

    def findings(self):
        peak=self.annual.loc[self.annual.registros.idxmax()]
        change=self.annual.loc[self.annual.variacion_pct.idxmax()]
        period=self.periods.set_index('periodo').promedio_anual
        male=self.sex.set_index('sexo_nombre').loc['Hombres']
        older=int(self.age.loc[self.age.grupo_edad.isin(['65–79','80+']),'registros'].sum())
        first=self.groups.iloc[0];leading=self.diseases.iloc[0]
        infant=self.age_leaders.loc[self.age_leaders.edad_amplia.eq('<1 año')].sort_values('registros',ascending=False).iloc[0]
        low=self.territory.loc[self.territory.cardiovascular_pct.idxmin()];high=self.territory.loc[self.territory.cardiovascular_pct.idxmax()]
        q=self.quality.loc[self.quality.indicador.eq('Residencia ignorada')]
        worst=q.loc[q.porcentaje.idxmax()]
        rows=[
            ('Cambio temporal observado',f'El máximo es {int(peak.anio)}, con {number(peak.registros)} registros; el mayor cambio interanual es {int(change.anio)} ({percent(change.variacion_pct)}).','Serie anual y variación interanual.','No hay línea base esperada; cobertura y registro pueden influir. No se identifica exceso ni causalidad.','¿Qué estabilidad temporal presentan los perfiles cuando se separan las perturbaciones del registro y la composición?'),
            ('Períodos de distinta duración',f'Los promedios anuales son {period[PERIODS[0]]:,.2f}, {period[PERIODS[1]]:,.2f} y {period[PERIODS[2]]:,.2f} registros para los tres períodos, respectivamente.','Tabla de períodos y gráfico de medias anuales.','Son intervalos operativos de 11, 2 y 3 años; las medias no controlan crecimiento ni estructura poblacional.','¿Persisten las diferencias de composición territorial al comparar ventanas de igual duración?'),
            ('Composición demográfica',f'Los hombres representan {percent(male.porcentaje)}; las edades de 65+ suman {percent(older/self.n*100)}.','Distribuciones demográficas, pirámide y composiciones anuales.','Se describen defunciones; faltan denominadores por sexo y edad, y existen edades ignoradas.','¿Cuánto de los futuros perfiles territoriales corresponde a diferencias de composición etaria?'),
            ('Concentración epidemiológica',f'{first.grupo_epidemiologico} concentra {number(first.registros)} registros ({percent(first.porcentaje)}).','Ranking de grupos y heatmaps anuales y por período.','Los grupos provienen del catálogo v2 y difieren en amplitud; la frecuencia no valida una taxonomía clínica.','¿Qué perfiles persisten al usar proporciones en vez de conteos y al revisar categorías de distinta amplitud?'),
            ('Enfermedades y edad',f'La etiqueta global principal es {leading.enfermedad_normalizada} ({number(leading.registros)}); entre menores de un año encabeza {infant.enfermedad_normalizada} ({number(infant.registros)}).','Top veinte y comparaciones por grandes grupos etarios.','Las etiquetas tienen especificidad desigual y algunas fueron reclasificadas manualmente. No prueban riesgo ni causalidad.','¿Qué patrones de causas se sostienen al separar edad, categorías residuales y códigos detallados?'),
            ('Heterogeneidad de composición territorial',f'La participación cardiovascular va de {percent(low.cardiovascular_pct)} en {low.departamento} a {percent(high.cardiovascular_pct)} en {high.departamento}.','Mapa de porcentaje cardiovascular y heatmap territorial.','El denominador son registros del departamento; no población ni tasas ajustadas. Extranjero e ignorado no se mapean.','¿Se mantienen esos contrastes después de considerar edad y calidad del registro en una futura evaluación de perfiles?'),
            ('Calidad del registro como dimensión',f'El mayor porcentaje anual de residencia ignorada es {percent(worst.porcentaje)} en {int(worst.anio)}; se conservan {number(self.ignored)} edades ignoradas y {number(self.nonspecific)} registros R99X.','Indicadores anuales de calidad y mapa de causas mal definidas.','La falta de especificidad puede generar perfiles de calidad antes que perfiles epidemiológicos; no se corrige arbitrariamente.','¿Qué estabilidad de agrupamiento territorial se obtiene al evaluar explícitamente la sensibilidad a desconocidos y causas mal definidas?'),
        ]
        for i,(title,pattern,evidence,limit,question) in enumerate(rows,1):
            display(Markdown(f'### {i}. {title}\n\n**Patrón:** {pattern}\n\n**Respaldo:** {evidence}\n\n**Límite:** {limit}\n\n**Pregunta posterior:** {question}'))

    def finish(self):
        assert all(sha256(p.read_bytes()).hexdigest()==h for p,h in self.protected.items())
        validation={'registros':self.n,'anios':self.annual.anio.astype(int).tolist(),
                    'total_anual':int(self.annual.registros.sum()),'total_grupos':int(self.groups.registros.sum()),
                    'total_enfermedades':int(self.diseases.registros.sum()),'grupos':len(self.groups),
                    'enfermedades':len(self.diseases),'poligonos':len(self.geo['features']),
                    'mapeados':int(self.territory.registros.sum()),'excluidos_mapa':int(self.excluded.registros.sum()),
                    'figuras':self.figures,'sha256_parquet':self.protected[self.path], 'archivos_fuente_intactos':True}
        display(HTML('<script id="eda-validacion" type="application/json">'+json.dumps(validation)+'</script>'))
        display(Markdown(f'Validaciones ejecutadas: totales anual, epidemiológico, por enfermedad, sexo y edad conciliados con **{number(self.n)}**; 19 grupos conservados; participaciones anuales y departamentales completas suman 100 %; 22 códigos y nombres geográficos verificados; excluidos geográficos reconciliados. **{self.figures} figuras** con datos agregados. Parquet, diccionarios y GeoJSON conservan sus huellas.'))
        display(Markdown(f'SHA-256 del Parquet leído: `{self.protected[self.path]}`. La ejecución cargó el Parquet una sola vez; las figuras reutilizan agregados en memoria. El HTML incluye las figuras, Plotly y polígonos, sin microdatos ni servidor Python posterior.'))
