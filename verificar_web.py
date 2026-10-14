"""Comprobación del panel con el ejecutor de pruebas de Streamlit."""
from pathlib import Path
import sys

raiz = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(raiz / "outputs"))
from streamlit.testing.v1 import AppTest

panel = AppTest.from_file(str(raiz / "outputs" / "app.py"), default_timeout=60).run()
assert not panel.exception, [e.message for e in panel.exception]
assert not panel.error, [e.value for e in panel.error]
assert len(panel.metric) == 4
assert panel.metric[1].value == "72.37%"
assert panel.metric[2].value == "63.75%"
assert len(panel.tabs) == 4
assert len(panel.dataframe[1].value) == 12
panel.multiselect[0].select("Cimentaciones").run()
assert not panel.exception
assert len(panel.dataframe[1].value) == 1
assert panel.dataframe[1].value.iloc[0]["Partida"] == "Cimentaciones"
panel.selectbox[0].select("Seguimiento").run()
assert not panel.exception
assert len(panel.dataframe[2].value) == 384
panel.button[0].click().run()
assert not panel.exception
assert panel.metric[2].value == "63.75%"
print("Panel verificado: 4 indicadores, 4 pestañas, filtro por partida, 384 registros y actualización de Excel.")
