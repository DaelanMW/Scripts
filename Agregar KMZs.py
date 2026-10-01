import os
import re
from qgis.PyQt.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QComboBox,
    QPushButton, QMessageBox, QScrollArea, QWidget, QFormLayout
)
from qgis.core import (
    QgsProject, QgsVectorLayer, QgsFeature,
    QgsCoordinateTransform
)

# -----------------------------
# CAPAS DESTINO
# -----------------------------
LAYER_NAMES = [
    "NO IMPORTAR",
    "Propiedad minera",
    "PM",
    "Limno",
    "Flora",
    "Monitoreo de Agua",
    "Estación meteorológica",
    "Caminos"
]

POLYGON_LAYER = "Propiedad minera"
LINE_LAYER = "Caminos"


# -----------------------------
# MATCH AUTOMÁTICO
# -----------------------------
def normalize(text):
    text = text.lower()
    text = re.sub(r'[_\-\.]', ' ', text)
    return text


def auto_match_layer(text):
    """text puede ser el nombre de archivo, el de una subcapa/folder, o la
    combinación de ambos (para mejorar la sugerencia)."""
    name = normalize(text)

    rules = {
        "Caminos": ["camino", "ruta", "road", "sendero"],
        "Monitoreo de Agua": ["agua", "aguas", "hidro", "hidrolog", "limno"],
        "Estación meteorológica": ["meteor", "clima", "weather", "estacion"],
        "Flora": ["flora", "vegetacion", "vegetación", "plant"],
        "Propiedad minera": ["mina", "minera", "propiedad"],
        "PM": ["pm", "punto", "puntos", "arqueo", "sitio", "sitios", "petroglifo", "relevamiento"],
    }

    for layer, keywords in rules.items():
        for kw in keywords:
            if kw in name:
                return layer

    return "PM"


# -----------------------------
# PROYECTO
# -----------------------------
project = QgsProject.instance()

if not project.fileName():
    raise Exception("Guardá el proyecto (.qgz) antes de ejecutar")

project_folder = os.path.dirname(project.fileName())
target_crs = project.crs()

# -----------------------------
# CARPETA KMZ
# -----------------------------
kmz_folder = os.path.join(project_folder, "KMZs")
search_folder = kmz_folder if os.path.exists(kmz_folder) else project_folder

files = [
    f for f in os.listdir(search_folder)
    if f.lower().endswith(".kmz") or f.lower().endswith(".kml")
]

if not files:
    raise Exception(f"No se encontraron KMZ/KML en: {search_folder}")


# -----------------------------
# DETECCIÓN DE SUBCAPAS (carpetas dentro del mismo KML/KMZ)
# -----------------------------
def probe_sublayers(path):
    """Devuelve una lista de dicts {id, name, count} por cada subcapa
    (carpeta) que el driver OGR detecta dentro del KML/KMZ.
    Si el archivo no tiene subcapas separables, devuelve una lista vacía
    y el archivo se trata como una única categoría."""
    probe = QgsVectorLayer(path, "probe", "ogr")

    sublayers = []
    raw_subs = probe.dataProvider().subLayers() if probe.isValid() else []

    if len(raw_subs) > 1:
        for sub in raw_subs:
            parts = sub.split("!!::!!")
            # formato típico: layerId!!::!!layerName!!::!!featureCount!!::!!wkbType!!::!!geomType
            layer_id = parts[0] if len(parts) > 0 else None
            layer_name = parts[1] if len(parts) > 1 else None
            count = parts[2] if len(parts) > 2 else "?"
            if layer_name:
                sublayers.append({"id": layer_id, "name": layer_name, "count": count})

    return sublayers


def open_sublayer(path, layer_name):
    """Abre una subcapa puntual (carpeta) de un KML/KMZ por nombre."""
    uri = f"{path}|layername={layer_name}"
    return QgsVectorLayer(uri, layer_name, "ogr")


# -----------------------------
# ARMADO DE FILAS (archivo [+ subcategoría opcional])
# -----------------------------
# Cada fila representa lo que se va a importar de manera independiente:
# - archivo simple -> 1 fila (subcategoria = None)
# - archivo con varias carpetas internas -> 1 fila por carpeta
rows_info = []  # lista de dicts: {file, sublayer_name_or_None, display_sub}

for f in files:
    path = os.path.join(search_folder, f)
    subs = probe_sublayers(path)

    if not subs:
        rows_info.append({"file": f, "sublayer": None, "display_sub": "-"})
    else:
        for s in subs:
            rows_info.append({
                "file": f,
                "sublayer": s["name"],
                "display_sub": f'{s["name"]} ({s["count"]} feat.)'
            })


# -----------------------------
# CARGA ROBUSTA (fallback para archivos sin subLayers detectables)
# -----------------------------
def load_whole_file(path, name):
    layers = []

    base = QgsVectorLayer(path, name, "ogr")

    if base.isValid() and base.featureCount() > 0:
        return [base]

    tmp = QgsVectorLayer(path, name, "ogr")

    if tmp.isValid():
        sublayers = tmp.dataProvider().subLayers()

        if sublayers:
            for sub in sublayers:
                parts = sub.split("!!::!!")
                layer_name = parts[1] if len(parts) > 1 else None
                if layer_name:
                    lyr = open_sublayer(path, layer_name)
                    if lyr.isValid():
                        layers.append(lyr)

    return layers


# -----------------------------
# DIALOGO
# -----------------------------
class ImportDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Importador KMZ/KML (Control total, con subcategorías)")

        layout = QVBoxLayout()
        form = QFormLayout()

        # cada fila guarda: (row_info, combo)
        self.row_widgets = []

        header = QLabel("<b>Archivo — Subcategoría</b>")
        form.addRow(header, QLabel("<b>Capa destino</b>"))

        for info in rows_info:
            label_text = info["file"]
            if info["sublayer"] is not None:
                label_text += f"  ›  {info['display_sub']}"

            label = QLabel(label_text)
            combo = QComboBox()
            combo.addItems(LAYER_NAMES)

            # sugerencia automática combinando nombre de archivo + subcapa
            match_text = info["file"] + " " + (info["sublayer"] or "")
            suggested = auto_match_layer(match_text)
            idx = combo.findText(suggested)
            if idx >= 0:
                combo.setCurrentIndex(idx)

            form.addRow(label, combo)
            self.row_widgets.append((info, combo))

        container = QWidget()
        container.setLayout(form)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(container)

        layout.addWidget(scroll)

        btn = QPushButton("Importar")
        btn.clicked.connect(self.run_import)

        layout.addWidget(btn)
        self.setLayout(layout)


    # -----------------------------
    # IMPORTACIÓN
    # -----------------------------
    def run_import(self):
        project = QgsProject.instance()
        total = 0

        # -------------------------
        # 1) Determinar capas destino usadas y limpiarlas UNA sola vez
        #    (varias filas/subcategorías pueden apuntar a la misma capa)
        # -------------------------
        target_layers_used = set(
            combo.currentText()
            for _, combo in self.row_widgets
            if combo.currentText() != "NO IMPORTAR"
        )

        target_cache = {}

        for layer_name in target_layers_used:
            target_list = project.mapLayersByName(layer_name)

            if not target_list:
                QMessageBox.warning(self, "Error", f"No existe capa: {layer_name}")
                continue

            target_layer = target_list[0]
            target_layer.startEditing()
            target_layer.dataProvider().deleteFeatures(
                [f.id() for f in target_layer.getFeatures()]
            )
            target_layer.commitChanges()

            target_cache[layer_name] = target_layer

        # -------------------------
        # 2) Procesar cada fila (archivo + subcategoría) por separado
        # -------------------------
        for info, combo in self.row_widgets:

            layer_name = combo.currentText()

            if layer_name == "NO IMPORTAR":
                continue

            target_layer = target_cache.get(layer_name)
            if target_layer is None:
                continue

            path = os.path.join(search_folder, info["file"])

            if info["sublayer"] is not None:
                src_layers = [open_sublayer(path, info["sublayer"])]
                src_layers = [l for l in src_layers if l.isValid()]
            else:
                src_layers = load_whole_file(path, info["file"])

            if not src_layers:
                QMessageBox.warning(
                    self, "Error",
                    f"No se pudo leer: {info['file']}"
                    + (f" / {info['sublayer']}" if info["sublayer"] else "")
                )
                continue

            transform = QgsCoordinateTransform(
                src_layers[0].crs(),
                target_crs,
                project
            )

            prov = target_layer.dataProvider()

            for src_layer in src_layers:

                for feat in src_layer.getFeatures():

                    geom = feat.geometry()
                    if not geom or geom.isEmpty():
                        continue

                    try:
                        geom.transform(transform)
                    except Exception:
                        continue

                    gtype = geom.type()

                    if layer_name == POLYGON_LAYER and gtype != 2:
                        continue
                    if layer_name == LINE_LAYER and gtype != 1:
                        continue
                    if layer_name not in [POLYGON_LAYER, LINE_LAYER] and gtype != 0:
                        continue

                    new_feat = QgsFeature()
                    new_feat.setGeometry(geom)
                    new_feat.setAttributes(feat.attributes())

                    prov.addFeature(new_feat)
                    total += 1

            target_layer.updateExtents()

        # -----------------------------
        # ZOOM FINAL
        # -----------------------------
        pm = project.mapLayersByName(POLYGON_LAYER)

        if pm and pm[0].featureCount() > 0:
            extent = pm[0].extent()
            extent.scale(1.1)
            iface.mapCanvas().setExtent(extent)
            iface.mapCanvas().refresh()

        QMessageBox.information(
            self,
            "OK",
            f"Importación finalizada\nFeatures: {total}"
        )

        self.close()


# -----------------------------
# EJECUTAR
# -----------------------------
dlg = ImportDialog()
dlg.exec_()