"""Panel local de control de obra. Ejecutar con streamlit run outputs/app.py."""
from hashlib import sha256
from io import BytesIO

import pandas as pd
import streamlit as st

from graficar_obra import BASE, graficar, leer_y_calcular, plt

plt.switch_backend("Agg")
st.set_page_config(page_title="Control de obra", page_icon="🏗️", layout="wide")


@st.cache_data(show_spinner="Leyendo el Excel y preparando los gráficos…")
def cargar_panel(contenido: bytes):
    """Recalcula tablas y gráficos cuando cambia el contenido del Excel."""
    parametros, corte, presupuesto, semanal, resumen = leer_y_calcular(BytesIO(contenido))
    carpeta = BASE.parent / "work" / "web" / sha256(contenido).hexdigest()[:16]
    try:
        graficar(parametros, corte, presupuesto, semanal, resumen, carpeta)
        imagenes = {nombre: (carpeta / f"{nombre}.png").read_bytes()
                    for nombre in ("curva_s", "avance_partidas", "valorizacion_semanal")}
        pdfs = {nombre: (carpeta / f"{nombre}.pdf").read_bytes() for nombre in imagenes}
    finally:
        plt.close("all")
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        semanal.to_excel(writer, sheet_name="Curva_S")
        resumen.to_excel(writer, sheet_name="Avance_partidas", index=False)
    hojas = pd.read_excel(BytesIO(contenido), sheet_name=None)
    return str(parametros["proyecto"]), corte, float(presupuesto), semanal, resumen, imagenes, pdfs, buffer.getvalue(), hojas


st.title("Control de avance de obra")
st.caption("Curva S y seguimiento por partidas · ejercicio con datos sintéticos")
st.markdown("""<style>
    [data-testid="stMetric"] {background:#171021; border:1px solid #40334F;
        border-radius:12px; padding:16px;}
    [data-testid="stMetricValue"] {color:#D8BFFF;}
    h1, h2, h3 {color:#D8BFFF;}
</style>""", unsafe_allow_html=True)

excel = BASE / "datos_obra.xlsx"
if not excel.exists():
    st.error("No se encontró datos_obra.xlsx. Ejecuta primero graficar_obra.py para generar el ejemplo.")
    st.stop()

with st.sidebar:
    st.header("Seguimiento de obra")
    st.caption("Fuente: datos_obra.xlsx")
    st.button("Actualizar desde Excel", help="Vuelve a leer el archivo después de guardar tus cambios.")
    st.caption("Los avances reales posteriores al corte permanecen sin datos.")

try:
    contenido = excel.read_bytes()
    proyecto, corte, presupuesto, semanal, resumen, imagenes, pdfs, archivo_resumen, hojas = cargar_panel(contenido)
except Exception as error:
    st.error(f"No se pudo cargar el panel. Revisa el Excel: {error}")
    st.stop()

st.subheader(proyecto)
st.caption(f"Corte de ejecución: semana {corte} de {len(semanal)} · {len(resumen)} partidas")
fila = semanal.loc[corte]
brecha = fila.real_pct - fila.plan_pct
columnas = st.columns(4)
columnas[0].metric("Presupuesto contractual", f"S/ {presupuesto:,.0f}")
columnas[1].metric("Avance planificado", f"{fila.plan_pct:.2f}%")
columnas[2].metric("Avance ejecutado", f"{fila.real_pct:.2f}%")
columnas[3].metric("Brecha de avance", f"{brecha:+.2f} pp")

with st.sidebar:
    st.divider()
    st.download_button("Descargar datos originales", contenido, "datos_obra.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("Descargar resumen calculado", archivo_resumen, "resumen_avance.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

curva, partidas, valorizacion, datos = st.tabs(["Curva S", "Avance por partidas", "Valorización semanal", "Datos de Excel"])


def mostrar_grafico(nombre, descripcion):
    st.image(imagenes[nombre], width="stretch")
    st.caption(descripcion)
    col1, col2 = st.columns(2)
    col1.download_button("Descargar PNG", imagenes[nombre], f"{nombre}.png", "image/png", key=f"png_{nombre}")
    col2.download_button("Descargar PDF", pdfs[nombre], f"{nombre}.pdf", "application/pdf", key=f"pdf_{nombre}")


with curva:
    mostrar_grafico("curva_s", "Avance acumulado ponderado por presupuesto. La ejecución real se muestra solo hasta la semana de corte.")
    with st.expander("Ver valores de la curva S"):
        st.dataframe(semanal.reset_index().round(2), hide_index=True, width="stretch")

with partidas:
    mostrar_grafico("avance_partidas", "Porcentaje de metrado acumulado respecto al metrado contractual de cada partida.")
    seleccion = st.multiselect("Consultar partidas", resumen.partida.tolist(), placeholder="Todas las partidas")
    tabla = resumen if not seleccion else resumen[resumen.partida.isin(seleccion)]
    tabla = tabla[["partida", "unidad", "cantidad_total", "cantidad_plan", "cantidad_real", "presupuesto", "plan_pct", "real_pct", "brecha_pp"]]
    st.dataframe(tabla.rename(columns={"partida": "Partida", "unidad": "Unidad", "cantidad_total": "Metrado contractual",
        "cantidad_plan": "Metrado planificado", "cantidad_real": "Metrado ejecutado", "presupuesto": "Presupuesto (S/)",
        "plan_pct": "Plan (%)", "real_pct": "Real (%)", "brecha_pp": "Brecha (pp)"}).round(2), hide_index=True, width="stretch")
    st.caption("El filtro permite consultar la tabla; el gráfico conserva las 12 partidas para comparar el conjunto.")

with valorizacion:
    mostrar_grafico("valorizacion_semanal", "Producción semanal valorizada a precios contractuales, expresada en miles de soles. No representa gastos ni pagos realizados.")

with datos:
    hoja = st.selectbox("Hoja del Excel", list(hojas))
    tabla_hoja = hojas[hoja].copy()
    if hoja == "Parametros":
        # La columna mezcla texto y números; mostrarla como texto evita
        # conversiones automáticas al formato de tablas del navegador.
        tabla_hoja["valor"] = tabla_hoja["valor"].astype("string")
    st.dataframe(tabla_hoja, hide_index=True, width="stretch")
    st.caption("Para actualizar los datos, edita y guarda el Excel y pulsa «Actualizar desde Excel».")

st.divider()
st.caption("Datos sintéticos para aprendizaje. Brecha = avance ejecutado − avance planificado, en puntos porcentuales.")
