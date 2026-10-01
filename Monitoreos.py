from qgis.core import QgsProject, QgsLayoutItemMap, QgsRectangle
from qgis.PyQt.QtWidgets import QMessageBox

LAYOUT_NAME = "Monitoreo"
MAP_ID = "Mapa 1"

project = QgsProject.instance()
layout = project.layoutManager().layoutByName(LAYOUT_NAME)

map_item = layout.itemById(MAP_ID)

def L(name):
    layers = project.mapLayersByName(name)
    if not layers:
        raise Exception(f"No existe la capa: {name}")
    return layers[0]

# =========================
# CAPAS
# =========================
google = L("Google")
rios = L("Ríos 1")
propiedad = L("Propiedad minera")
rios_etiq = L("Rios etiquetas")
sm_pm = L("PM")
sm_limno = L("Limno")
sm_flora = L("Flora")
sm_agua = L("Monitoreo de Agua")
sm_meteo = L("Estación meteorológica")


# =========================
# EXTENT BASE
# =========================
extent = propiedad.extent()

if extent.isEmpty():
    raise Exception("Extent vacío en 'Propiedad minera'")

# =========================
# VALIDACIÓN MAP ITEM
# =========================
if not isinstance(map_item, QgsLayoutItemMap):
    raise Exception("El item 'Mapa 1' no es un mapa")

# =========================
# 🔥 ORDEN DE CAPAS (CRÍTICO)
# =========================
map_item.setKeepLayerSet(False)

map_item.setLayers([
    sm_pm,
    sm_limno,
    sm_flora,
    sm_agua,
    sm_meteo,
    rios,
    rios_etiq,
    propiedad,
    google
])

# =====================================================
# 🔥 EXPANDIR 10% (ZOOM OUT)
# =====================================================
# Corregido: Se usa 'extent' en lugar de 'ext'
dx = extent.width() * 0.05
dy = extent.height() * 0.05

expanded = QgsRectangle(
    extent.xMinimum() - dx,
    extent.yMinimum() - dy,
    extent.xMaximum() + dx,
    extent.yMaximum() + dy
)

# =====================================================
# APLICAR AL MAPA (SIN CAMBIAR TAMAÑO)
# =====================================================
map_item.zoomToExtent(expanded)

# REFRESH LIMPIO
map_item.refresh()
layout.refresh()

# ABRIR DISEÑADOR AL FINAL
iface.openLayoutDesigner(layout)
