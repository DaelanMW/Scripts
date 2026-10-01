import math
from qgis.core import QgsProject

project = QgsProject.instance()

# =========================
# REDONDEO CARTOGRÁFICO
# =========================
def round_scale(s):
    if s <= 0:
        return s

    magnitude = 10 ** int(math.floor(math.log10(s)))
    normalized = s / magnitude

    steps = [1, 1.25, 1.5, 1.75, 2,
             2.5, 3, 4, 5, 6, 7.5, 8, 9, 10]

    closest = min(steps, key=lambda x: abs(x - normalized))
    return int(closest * magnitude)

# =========================
# RECORRER TODAS LAS COMPOSICIONES
# =========================
layouts = project.layoutManager().printLayouts()

if not layouts:
    raise Exception("No hay composiciones en el proyecto")

for layout in layouts:

    map_item = layout.itemById("Mapa 1")
    label_item = layout.itemById("Escala 2")

    if not map_item or not label_item:
        continue  # saltar layouts incompletos

    scale = map_item.scale()
    rounded = round_scale(scale)

    label_item.setText(f"Escala 1:{rounded}")

    label_item.refresh()
    layout.refresh()

print("Escalas actualizadas en todas las composiciones")