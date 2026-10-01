from qgis.PyQt.QtWidgets import QInputDialog, QMessageBox
from qgis.core import (
    QgsProject,
    QgsLayoutItemMap,
    QgsVectorLayer,
    QgsSingleSymbolRenderer,
    QgsFillSymbol
)
from datetime import datetime

# =====================================================
# CONFIG
# =====================================================

ID_TITULO = "Titulo"
ID_FECHA = "Fecha"
MAP_ID = "Localidad"

TEMP_LAYER_NAME = "Propiedad minera roja"

MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

# =====================================================
# INPUT
# =====================================================

texto, ok = QInputDialog.getText(
    None,
    "Actualizar layouts",
    "Nuevo TÍTULO:"
)

if not ok or not texto:
    raise Exception("Cancelado")

# =====================================================
# FECHA
# =====================================================

now = datetime.now()
fecha_texto = f"FECHA: {MESES[now.month]} {now.year}".upper()

# =====================================================
# PROYECTO
# =====================================================

project = QgsProject.instance()
manager = project.layoutManager()

layouts = manager.printLayouts()

# =====================================================
# CAPAS
# =====================================================

def capa(nombre):
    capas = project.mapLayersByName(nombre)
    if not capas:
        raise Exception(f"No existe la capa: {nombre}")
    return capas[0]

google = capa("Google")
localidades = capa("Localidades")
referencias = capa("referencias")
propiedad = capa("Propiedad minera")

# =====================================================
# CAPA TEMPORAL ROJA
# =====================================================

capas_rojas = project.mapLayersByName(TEMP_LAYER_NAME)

if capas_rojas:
    propiedad_roja = capas_rojas[0]
else:
    propiedad_roja = QgsVectorLayer(
        propiedad.source(),
        TEMP_LAYER_NAME,
        propiedad.providerType()
    )

    if not propiedad_roja.isValid():
        raise Exception("No se pudo crear la capa temporal")

    symbol = QgsFillSymbol.createSimple({
        'color': '255,0,0,80',
        'outline_color': '255,0,0',
        'outline_width': '0.8'
    })

    propiedad_roja.setRenderer(
        QgsSingleSymbolRenderer(symbol)
    )

    project.addMapLayer(propiedad_roja, False)

# =====================================================
# RESULTADOS
# =====================================================

layouts_ok = []
layouts_sin_titulo_fecha = []
layouts_sin_localidad = []

# =====================================================
# LOOP TODOS LOS LAYOUTS
# =====================================================

for layout in layouts:

    nombre_layout = layout.name()

    # ---------------------------------
    # TITULO / FECHA
    # ---------------------------------

    titulo = layout.itemById(ID_TITULO)
    fecha = layout.itemById(ID_FECHA)

    if titulo and hasattr(titulo, "setText"):
        titulo.setText(texto)

    if fecha and hasattr(fecha, "setText"):
        fecha.setText(fecha_texto)

    if not titulo or not fecha:
        layouts_sin_titulo_fecha.append(nombre_layout)

    # ---------------------------------
    # MAPA LOCALIDAD
    # ---------------------------------

    mapa_localidad = layout.itemById(MAP_ID)

    if (
        mapa_localidad and
        isinstance(mapa_localidad, QgsLayoutItemMap)
    ):

        mapa_localidad.setKeepLayerSet(True)

        mapa_localidad.setLayers([
            propiedad_roja,   # arriba
            referencias,
            localidades,
            google            # fondo
        ])

        mapa_localidad.refresh()

    else:
        layouts_sin_localidad.append(nombre_layout)

    layout.refresh()
    layouts_ok.append(nombre_layout)

# =====================================================
# RESUMEN
# =====================================================

msg = "RESULTADO\n\n"

msg += f"Layouts procesados: {len(layouts_ok)}\n\n"

if layouts_sin_titulo_fecha:
    msg += (
        "Sin items Titulo/Fecha:\n" +
        "\n".join(layouts_sin_titulo_fecha) +
        "\n\n"
    )

if layouts_sin_localidad:
    msg += (
        "Sin mapa Localidad:\n" +
        "\n".join(layouts_sin_localidad)
    )

QMessageBox.information(
    None,
    "Layouts actualizados",
    msg
)

print("Proceso finalizado correctamente")