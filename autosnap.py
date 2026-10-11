# -*- coding: utf-8 -*-
"""SLS Magic AutoSnap Bangunan v2.8.5 - georeference selector matches SLS group names - menu UI and Excel audit export, code-only village selector, dependent SLS, filtered map, counts and enriched tables."""
import os, math, random, re, tempfile, shutil, json, urllib.request, urllib.error, webbrowser
try:
    from .village_reference import VILLAGE_BY_CODE
except Exception:
    VILLAGE_BY_CODE = {}
from qgis.PyQt.QtCore import Qt
from qgis.PyQt.QtGui import QColor, QFont, QIcon, QPainter, QPen, QImage
from qgis.PyQt.QtWidgets import (QAction, QMenu, QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QTableWidget, QTableWidgetItem, QFileDialog, QMessageBox, QCheckBox, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QAbstractItemView, QTabWidget, QSpinBox, QDoubleSpinBox, QProgressBar, QSplitter, QFrame, QSizePolicy, QWidget, QLineEdit, QInputDialog, QToolTip, QStyle, QToolButton)
from qgis.gui import QgsMapToolEmitPoint
from qgis.core import (QgsProject, QgsWkbTypes, QgsFeature, QgsGeometry, QgsPointXY,
    QgsCoordinateTransform, QgsVectorLayer, QgsVectorFileWriter, QgsPalLayerSettings,
    QgsVectorLayerSimpleLabeling, QgsTextFormat, QgsTextBufferSettings,
    QgsTextShadowSettings, QgsMarkerSymbol, QgsSingleSymbolRenderer, QgsUnitTypes,
    QgsRasterLayer, QgsPoint, Qgis, QgsRectangle, QgsRuleBasedRenderer, QgsField, QgsCategorizedSymbolRenderer, QgsRendererCategory, QgsFeatureRequest)
from qgis.PyQt.QtCore import QVariant


def _configure_tesseract_ocr(pytesseract_module):
    """Locate Tesseract executable for QGIS Python and verify it actually runs."""
    candidates = []
    # Respect an explicit user/system configuration first.
    env_cmd = os.environ.get('TESSERACT_CMD', '').strip().strip('"')
    if env_cmd:
        candidates.append(env_cmd)
    path_cmd = shutil.which('tesseract')
    if path_cmd:
        candidates.append(path_cmd)
    # Common Windows installations used alongside QGIS/OSGeo4W.
    candidates.extend([
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        r'C:\OSGeo4W\bin\tesseract.exe',
        r'C:\OSGeo4W64\bin\tesseract.exe',
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Tesseract-OCR', 'tesseract.exe'),
    ])
    seen = set()
    errors = []
    for candidate in candidates:
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        if not os.path.isfile(candidate):
            continue
        try:
            pytesseract_module.pytesseract.tesseract_cmd = candidate
            version = pytesseract_module.get_tesseract_version()
            return candidate, str(version).splitlines()[0]
        except Exception as exc:
            errors.append(f'{candidate}: {exc}')
    # Last try lets pytesseract resolve the executable through PATH on its own.
    try:
        pytesseract_module.pytesseract.tesseract_cmd = 'tesseract'
        version = pytesseract_module.get_tesseract_version()
        return 'tesseract (PATH)', str(version).splitlines()[0]
    except Exception as exc:
        errors.append(str(exc))
    raise RuntimeError('Executable Tesseract tidak ditemukan atau tidak dapat dijalankan. '
                       'Instal Tesseract OCR, lalu pastikan tesseract.exe ada di PATH atau '
                       'atur variabel lingkungan TESSERACT_CMD. ' + ' | '.join(errors[-3:]))



class GeoreferenceConvertDialog(QDialog):
    """Convert georeferenced PNG/worldfile/aux XML or rasters in a selected QGIS group to GeoTIFF."""
    def __init__(self, iface, parent=None):
        super().__init__(parent)
        self.iface = iface
        self.setWindowTitle('Toolbox | Konversi Georeference ke GeoTIFF')
        self.resize(720, 360)
        self.setMinimumSize(620, 300)
        root = QVBoxLayout(self)
        title = QLabel('KONVERSI GEOREFERENCE → GEOTIFF')
        title.setStyleSheet('font-size:16pt;font-weight:800;color:#12304a;')
        root.addWidget(title)
        info = QLabel('PNG akan dibaca bersama worldfile PGW dan/atau metadata XML pendamping jika tersedia. Untuk ekspor grup, setiap raster di dalam grup dibuat menjadi file TIF tersendiri.')
        info.setWordWrap(True)
        info.setStyleSheet('color:#607086;')
        root.addWidget(info)
        self.tabs = QTabWidget()
        root.addWidget(self.tabs, 1)

        file_tab = QWidget()
        fl = QVBoxLayout(file_tab)
        fl.addWidget(QLabel('1. Pilih file PNG, PGW, XML, atau AUX.XML. File pendamping harus memiliki nama dasar yang sama dengan PNG.'))
        fr = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText('Pilih PNG atau file georeference pendamping…')
        fr.addWidget(self.file_edit, 1)
        browse = QPushButton('Pilih file…')
        browse.clicked.connect(self.browse_file)
        fr.addWidget(browse)
        fl.addLayout(fr)
        fl.addWidget(QLabel('Output GeoTIFF:'))
        outrow = QHBoxLayout()
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText('Pilih lokasi file .tif')
        outrow.addWidget(self.output_edit, 1)
        outbrowse = QPushButton('Simpan sebagai…')
        outbrowse.clicked.connect(self.browse_output)
        outrow.addWidget(outbrowse)
        fl.addLayout(outrow)
        fl.addStretch(1)
        self.tabs.addTab(file_tab, 'File PNG + PGW/XML')

        group_tab = QWidget()
        gl = QVBoxLayout(group_tab)
        gl.addWidget(QLabel('2. Pilih grup layer QGIS. Semua raster valid di dalam grup (termasuk subgrup) akan diekspor satu per satu ke GeoTIFF.'))
        self.group_combo = QComboBox()
        self.group_nodes = []
        self._collect_groups(QgsProject.instance().layerTreeRoot(), '')
        # Prefer the group currently selected in QGIS, when it is available.
        try:
            current_node = iface.layerTreeView().currentNode()
            for i, (path, node) in enumerate(self.group_nodes):
                if node == current_node:
                    self.group_combo.setCurrentIndex(i)
                    break
        except Exception:
            pass
        gl.addWidget(self.group_combo)
        gr = QHBoxLayout()
        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText('Folder tujuan ekspor TIF')
        gr.addWidget(self.folder_edit, 1)
        folder_btn = QPushButton('Pilih folder…')
        folder_btn.clicked.connect(self.browse_folder)
        gr.addWidget(folder_btn)
        gl.addLayout(gr)
        gl.addStretch(1)
        self.tabs.addTab(group_tab, 'Ekspor grup layer aktif')

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        self.run_btn = QPushButton('Konversi / Ekspor TIF')
        self.run_btn.setMinimumHeight(38)
        self.run_btn.setStyleSheet('QPushButton{background:#1769aa;color:white;font-weight:700;border:0;border-radius:6px;padding:8px 16px;} QPushButton:hover{background:#0f548b;}')
        self.run_btn.clicked.connect(self.run_conversion)
        buttons.addWidget(self.run_btn)
        close_btn = QPushButton('Tutup')
        close_btn.setMinimumHeight(38)
        close_btn.clicked.connect(self.close)
        buttons.addWidget(close_btn)
        root.addLayout(buttons)

    def _collect_groups(self, node, prefix):
        for child in node.children():
            if child.nodeType() == 0:
                name = str(child.name())
                path = (prefix + ' / ' + name) if prefix else name
                self.group_nodes.append((path, child))
                self.group_combo.addItem(path, len(self.group_nodes)-1)
                self._collect_groups(child, path)

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Pilih PNG atau file georeference',
            '', 'Georeference (*.png *.pgw *.wld *.xml *.aux.xml);;Semua file (*.*)')
        if path:
            self.file_edit.setText(path)
            low = path.lower()
            if low.endswith('.png'):
                base = os.path.splitext(path)[0]
            elif low.endswith('.aux.xml'):
                base = path[:-8]
            else:
                base = os.path.splitext(path)[0]
            self.output_edit.setText(base + '_georeferenced.tif')

    def browse_output(self):
        path, _ = QFileDialog.getSaveFileName(self, 'Simpan GeoTIFF', self.output_edit.text() or '', 'GeoTIFF (*.tif *.tiff)')
        if path:
            if not path.lower().endswith(('.tif', '.tiff')):
                path += '.tif'
            self.output_edit.setText(path)

    def browse_folder(self):
        path = QFileDialog.getExistingDirectory(self, 'Pilih folder output GeoTIFF', self.folder_edit.text() or '')
        if path:
            self.folder_edit.setText(path)

    def _resolve_png(self, chosen):
        low = chosen.lower()
        if low.endswith('.png'):
            return chosen
        if low.endswith('.aux.xml'):
            base = chosen[:-8]
        else:
            base = os.path.splitext(chosen)[0]
        for ext in ('.png', '.PNG'):
            candidate = base + ext
            if os.path.exists(candidate):
                return candidate
        return None

    def _export_dataset(self, source, destination, gdal):
        ds = gdal.Open(source)
        if ds is None:
            raise RuntimeError('GDAL tidak dapat membuka sumber: ' + source)
        try:
            options = gdal.TranslateOptions(format='GTiff', creationOptions=['COMPRESS=LZW', 'TILED=YES'])
            out = gdal.Translate(destination, ds, options=options)
            if out is None:
                raise RuntimeError('GDAL gagal menulis GeoTIFF: ' + destination)
            out.FlushCache()
            out = None
        finally:
            ds = None
        layer = QgsRasterLayer(destination, os.path.splitext(os.path.basename(destination))[0])
        if layer.isValid():
            QgsProject.instance().addMapLayer(layer)
            return True, layer.crs().authid()
        return False, ''

    def run_conversion(self):
        try:
            from osgeo import gdal
        except Exception as e:
            QMessageBox.critical(self, 'GDAL tidak tersedia', f'GDAL Python tidak dapat dimuat dari QGIS: {e}')
            return
        if self.tabs.currentIndex() == 0:
            chosen = self.file_edit.text().strip()
            output = self.output_edit.text().strip()
            if not chosen or not output:
                QMessageBox.information(self, 'Lengkapi pilihan', 'Pilih file sumber dan lokasi output GeoTIFF.')
                return
            png = self._resolve_png(chosen)
            if not png:
                QMessageBox.warning(self, 'PNG tidak ditemukan', 'Untuk input PGW/XML, file PNG dengan nama dasar yang sama harus berada di folder yang sama.')
                return
            # GDAL auto-loads a same-basename .pgw and .aux.xml when opening the PNG.
            sidecars = [os.path.splitext(png)[0] + ext for ext in ('.pgw', '.PGW', '.wld', '.WLD', '.aux.xml')]
            if not any(os.path.exists(s) for s in sidecars) and not chosen.lower().endswith('.xml'):
                answer = QMessageBox.question(self, 'File georeference tidak ditemukan',
                    'Tidak ditemukan PGW/WLD/AUX.XML pendamping. GeoTIFF mungkin tidak memiliki koordinat georeferensi. Lanjutkan?',
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if answer != QMessageBox.Yes:
                    return
            temporary_aux = None
            try:
                # GDAL recognizes PAM georeferencing metadata as <image>.aux.xml.
                # If the user explicitly selected a generic XML companion, make a temporary
                # recognized copy only when a proper .aux.xml does not already exist.
                if chosen.lower().endswith('.xml') and not chosen.lower().endswith('.aux.xml'):
                    expected_aux = os.path.splitext(png)[0] + '.aux.xml'
                    if not os.path.exists(expected_aux):
                        temporary_aux = expected_aux
                        shutil.copy2(chosen, temporary_aux)
                ok, crs = self._export_dataset(png, output, gdal)
                QMessageBox.information(self, 'Konversi selesai',
                    f'GeoTIFF berhasil dibuat:\n{output}\nCRS terdeteksi: {crs or "tidak ada"}\n\nPeriksa CRS dan posisi layer sebelum digunakan.')
            except Exception as e:
                QMessageBox.critical(self, 'Konversi gagal', f'{type(e).__name__}: {e}')
            finally:
                if temporary_aux and os.path.exists(temporary_aux):
                    try: os.remove(temporary_aux)
                    except Exception: pass
            return

        idx = self.group_combo.currentData()
        folder = self.folder_edit.text().strip()
        if idx is None or not folder:
            QMessageBox.information(self, 'Lengkapi pilihan', 'Pilih grup layer dan folder tujuan ekspor.')
            return
        try:
            os.makedirs(folder, exist_ok=True)
            group = self.group_nodes[int(idx)][1]
            layers = []
            def collect(node):
                for child in node.children():
                    if child.nodeType() == 0:
                        collect(child)
                    elif child.nodeType() == 1:
                        lyr = child.layer()
                        if isinstance(lyr, QgsRasterLayer) and lyr.isValid():
                            layers.append(lyr)
            collect(group)
            if not layers:
                QMessageBox.warning(self, 'Tidak ada raster', 'Grup terpilih tidak memiliki layer raster valid untuk diekspor.')
                return
            successes, errors = [], []
            for layer in layers:
                src = layer.source().split('|')[0]
                stem = re.sub(r'[^A-Za-z0-9_-]+', '_', layer.name()).strip('_') or 'raster'
                dest = os.path.join(folder, stem + '.tif')
                if os.path.abspath(src) == os.path.abspath(dest):
                    dest = os.path.join(folder, stem + '_converted.tif')
                try:
                    self._export_dataset(src, dest, gdal)
                    successes.append(dest)
                except Exception as e:
                    errors.append(f'{layer.name()}: {e}')
            message = f'Berhasil ekspor {len(successes)} dari {len(layers)} raster.'
            if successes:
                message += '\n\n' + '\n'.join(successes[:8])
            if errors:
                message += '\n\nGagal:\n' + '\n'.join(errors[:8])
            QMessageBox.information(self, 'Ekspor grup selesai', message)
        except Exception as e:
            QMessageBox.critical(self, 'Ekspor grup gagal', f'{type(e).__name__}: {e}')


class CircularProgress(QWidget):
    """Compact circular progress ring with a lightning glyph, compatible with QGIS PyQt."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._value = 0
        self.setMinimumSize(76, 76)
        self.setMaximumSize(96, 96)
        self.setToolTip('Progres pekerjaan AutoSnap')
    def setValue(self, value):
        self._value = max(0, min(100, int(value)))
        self.update()
    def value(self):
        return self._value
    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing, True)
        side = min(self.width(), self.height()) - 12
        x = (self.width()-side)/2; y = (self.height()-side)/2
        rect = __import__('qgis.PyQt.QtCore', fromlist=['QRectF']).QRectF(x, y, side, side)
        p.setPen(QPen(QColor('#dce7e1'), 7, Qt.SolidLine, Qt.RoundCap)); p.drawArc(rect, 0, 360*16)
        p.setPen(QPen(QColor('#22c55e'), 7, Qt.SolidLine, Qt.RoundCap)); p.drawArc(rect, 90*16, -int(360*16*self._value/100))
        p.setPen(QColor('#14532d')); p.setFont(QFont('Arial', 15, QFont.Bold)); p.drawText(rect, Qt.AlignCenter, '⚡')
        p.setFont(QFont('Arial', 8, QFont.Bold)); p.drawText(0, self.height()-2, self.width(), 12, Qt.AlignCenter, f'{self._value}%')
        p.end()

class AutoSnapDialog(QDialog):
    def __init__(self, iface, parent=None):
        super().__init__(parent)
        self.iface = iface
        self.setWindowTitle('SLS MAGIC  |  AutoSnap Bangunan  •  v2.8.5')
        self.resize(1120, 760)
        self.setMinimumSize(680, 460)
        self.setSizeGripEnabled(True)
        self.setWindowFlags(self.windowFlags() | Qt.Window)
        self.source = None
        self.mapping = {}       # source feature id -> target point in source CRS
        self._old_tool = None
        self._click_tool = None
        self._active_fid = None
        self._auto_candidates = {}  # fid -> suggested target
        self._preview_layer = None
        self._changed_layer = None
        self._pick_table = None
        self._raster_combo = None
        self._metadata_fields = {}
        self._feature_group_values = {}
        self._feature_records = {}
        self._row_lookup = {}
        self._layer_signature = None
        self._building_tables = False
        self._group_village_field = -1
        self._group_sls_field = -1
        self._group_uses_level6_as_primary = False
        self._original_subset_strings = {}
        self._subset_filter_applied = False
        self._georef_original_visibility = {}
        self._georef_managed_ids = set()
        self._selected_village_code = None
        self._selected_sls_name = None
        self._build_ui()
        self.refresh_layers()

    def _build_ui(self):
        root = QVBoxLayout(self); root.setContentsMargins(12,12,12,12); root.setSpacing(9)
        # Gunakan tombol jendela bawaan OS/QGIS di title bar; jangan gandakan
        # tombol minimize, maximize, dan close di dalam dialog.
        title = QLabel('<b style="font-size:17px">✦  SLS MAGIC  /  AUTOSNAP BANGUNAN</b>'); root.addWidget(title)
        subtitle = QLabel('Pilih kode desa, lalu nama SLS. Peta dan tabel hanya menampilkan titik pada SLS aktif; klik noktah hitam pada raster untuk memindahkan titik.')
        subtitle.setWordWrap(True); root.addWidget(subtitle)
        bar = QHBoxLayout(); bar.addWidget(QLabel('◉ Layer titik:'))
        self.layer_combo = QComboBox(); self.layer_combo.setMinimumWidth(150); self.layer_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.layer_combo.currentIndexChanged.connect(self.load_features)
        bar.addWidget(self.layer_combo, 1)
        self.refresh_btn = QPushButton('⟳ Muat ulang'); self.refresh_btn.clicked.connect(lambda: self.refresh_layers(force=True)); bar.addWidget(self.refresh_btn)
        root.addLayout(bar)

        # Hidden metadata fallbacks retained for output attributes; the duplicate text-entry row
        # is intentionally removed. Visible metadata is represented by the two clean selectors below.
        self.desa_edit = QLineEdit(self); self.desa_edit.hide()
        self.sls_edit = QLineEdit(self); self.sls_edit.hide()

        # Hierarchical work grouping: village first, then SLS name. Field names stay internal.
        group_bar = QFrame(); group_bar.setObjectName('metaPanel')
        gbl = QHBoxLayout(group_bar); gbl.setContentsMargins(12,9,12,9); gbl.setSpacing(10)
        self.group_primary_label = QLabel('🏘 Desa')
        gbl.addWidget(self.group_primary_label)
        self.group_primary_combo = QComboBox(); self.group_primary_combo.setMinimumWidth(175); self.group_primary_combo.setMaxVisibleItems(25)
        self.group_primary_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.group_primary_combo.currentIndexChanged.connect(self._on_primary_group_changed)
        gbl.addWidget(self.group_primary_combo, 2)
        self.group_sls_label = QLabel('▦ SLS')
        gbl.addWidget(self.group_sls_label)
        self.group_sls_combo = QComboBox(); self.group_sls_combo.setMinimumWidth(175); self.group_sls_combo.setMaxVisibleItems(25)
        self.group_sls_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.group_sls_combo.currentIndexChanged.connect(self._apply_group_filter)
        gbl.addWidget(self.group_sls_combo, 2)
        self.raster_sls_label = QLabel('▧ Raster SLS')
        gbl.addWidget(self.raster_sls_label)
        self.raster_sls_combo = QComboBox(); self.raster_sls_combo.setMinimumWidth(205); self.raster_sls_combo.setMaxVisibleItems(30); self.raster_sls_combo.setEnabled(True)
        self.raster_sls_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.raster_sls_combo.setToolTip('Pilih layer dari grup Georeference; layer terpilih akan dicentang dan diaktifkan.')
        self.raster_sls_combo.currentIndexChanged.connect(self._on_raster_sls_selected)
        gbl.addWidget(self.raster_sls_combo, 2)
        self.raster_sls_choose_btn = QPushButton('Choose')
        self.raster_sls_choose_btn.setToolTip('Pilih raster dari grup Georeference dan aktifkan layer tersebut di panel QGIS')
        self.raster_sls_choose_btn.clicked.connect(lambda: self._choose_georeference_raster('sls'))
        gbl.addWidget(self.raster_sls_choose_btn)
        root.addWidget(group_bar)

        # REAL UI REWORK: compact command bar inside the main dialog.
        # The existing tabs remain as the functional container, but are switched by popup
        # actions so controls do not stack vertically or duplicate on screen.
        command_bar = QFrame()
        command_bar.setObjectName('commandBar')
        cbl = QHBoxLayout(command_bar)
        cbl.setContentsMargins(8, 6, 8, 6)
        cbl.setSpacing(8)

        self.geotag_popup_btn = QToolButton()
        self.geotag_popup_btn.setText('Auto Geotagging  ▾')
        self.geotag_popup_btn.setPopupMode(QToolButton.InstantPopup)
        self.geotag_popup_menu = QMenu(self.geotag_popup_btn)
        self.geotag_manual_action = self.geotag_popup_menu.addAction('Manual')
        self.geotag_manual_action.triggered.connect(lambda checked=False: self.tabs.setCurrentIndex(0))
        self.geotag_auto_action = self.geotag_popup_menu.addAction('Otomatis')
        self.geotag_auto_action.triggered.connect(lambda checked=False: self.tabs.setCurrentIndex(1))
        self.geotag_popup_btn.setMenu(self.geotag_popup_menu)
        cbl.addWidget(self.geotag_popup_btn)

        self.toolbox_popup_btn = QToolButton()
        self.toolbox_popup_btn.setText('Toolbox  ▾')
        self.toolbox_popup_btn.setPopupMode(QToolButton.InstantPopup)
        self.toolbox_popup_menu = QMenu(self.toolbox_popup_btn)
        self.toolbox_style_action = self.toolbox_popup_menu.addAction('Style Layer')
        self.toolbox_style_action.triggered.connect(self.load_style_current)
        self.toolbox_load_style_action = self.toolbox_popup_menu.addAction('Load Style — Titik Bangunan')
        self.toolbox_load_style_action.triggered.connect(self._load_style_from_popup)
        self.toolbox_popup_menu.addSeparator()
        self.toolbox_convert_georef_action = self.toolbox_popup_menu.addAction('Konversi Georeference PNG/PGW/XML → GeoTIFF')
        self.toolbox_convert_georef_action.triggered.connect(self.open_georeference_converter)
        self.toolbox_popup_btn.setMenu(self.toolbox_popup_menu)
        cbl.addWidget(self.toolbox_popup_btn)

        self.audit_popup_btn = QPushButton('Audit Laporan Titik')
        self.audit_popup_btn.setToolTip('Buka laporan audit titik tanpa menumpuk panel di jendela utama')
        self.audit_popup_btn.clicked.connect(self._open_audit_from_dialog)
        cbl.addWidget(self.audit_popup_btn)

        self.note_popup_btn = QPushButton('Catatan Titik Bangunan')
        self.note_popup_btn.setToolTip('Pilih titik bermasalah pada layer titik aktif, lalu masukkan catatan')
        self.note_popup_btn.clicked.connect(self.start_note_pick)
        cbl.addWidget(self.note_popup_btn)

        cbl.addStretch(1)
        root.addWidget(command_bar)

        split = QSplitter(Qt.Horizontal); split.setChildrenCollapsible(False)
        left = QFrame(); left.setObjectName('leftPanel'); ll=QVBoxLayout(left); ll.setContentsMargins(0,0,6,0); ll.setSpacing(6)
        self.tabs = QTabWidget()
        # Tabs must be ordinary child widgets, not top-level QDialogs. Using QDialog here
        # can detach tab pages from the main window on some QGIS/PyQt builds.
        self.manual_tab = QWidget(); self.manual_tab.setLayout(QVBoxLayout())
        self.auto_tab = QWidget(); self.auto_tab.setLayout(QVBoxLayout())
        self.tabs.addTab(self.manual_tab, '①  Manual')
        self.tabs.addTab(self.auto_tab, '②  Otomatis')
        self.tabs.tabBar().hide()
        ll.addWidget(self.tabs,1); split.addWidget(left)

        right = QFrame(); right.setObjectName('rightPanel'); right.setMinimumWidth(220); right.setMaximumWidth(360)
        rl=QVBoxLayout(right); rl.setContentsMargins(12,12,12,12); rl.setSpacing(9)
        rl.addWidget(QLabel('<b style="font-size:14px">⚡ WORKFLOW & PROGRES</b>'))
        self.circular_progress = CircularProgress(); rl.addWidget(self.circular_progress, 0, Qt.AlignHCenter)
        self.current_target_label=QLabel('—')
        self.current_target_label.setAlignment(Qt.AlignCenter); self.current_target_label.setWordWrap(True); self.current_target_label.setObjectName('targetCard'); self.current_target_label.setMinimumHeight(74); rl.addWidget(self.current_target_label)
        self.next_number_caption=QLabel('NOMOR BERIKUTNYA'); self.next_number_caption.setAlignment(Qt.AlignCenter); self.next_number_caption.setStyleSheet('font-size:10px;font-weight:700;color:#64748b;letter-spacing:1px;'); rl.addWidget(self.next_number_caption)
        self.next_target_btn=QPushButton('—')
        self.next_target_btn.setObjectName('nextNumber'); self.next_target_btn.setMinimumHeight(58); self.next_target_btn.clicked.connect(self.go_next_target); rl.addWidget(self.next_target_btn)
        rl.addWidget(QLabel('<b>Progres pemetaan</b>'))
        self.process_progress=QProgressBar(); self.process_progress.setRange(0,100); self.process_progress.setValue(0); self.process_progress.setFormat('PEMETAAN • %p%'); rl.addWidget(self.process_progress)
        self.auto_progress=QProgressBar(); self.auto_progress.setRange(0,100); self.auto_progress.setValue(0); self.auto_progress.setFormat('PEMINDAIAN RASTER • %p%'); rl.addWidget(self.auto_progress)
        self.total_points_btn=QPushButton('TOTAL TITIK SLS AKTIF  •  0')
        self.total_points_btn.setMinimumHeight(38)
        self.total_points_btn.setToolTip('Jumlah titik yang cocok dengan kode desa dan SLS terpilih')
        self.total_points_btn.clicked.connect(self.show_active_group_count)
        rl.addWidget(self.total_points_btn)
        self.status_label=QLabel('Siap. Layer sumber tidak akan diubah.'); self.status_label.setWordWrap(True); self.status_label.setObjectName('statusCard'); rl.addWidget(self.status_label)
        rl.addStretch(1)
        self.review_btn=QPushButton('▣  REVIEW PETA')
        self.review_btn.setMinimumHeight(40); self.review_btn.clicked.connect(self.open_review); rl.addWidget(self.review_btn)
        self.finish_btn=QPushButton('✓  SELESAI • BUAT LAYER BARU')
        self.finish_btn.setMinimumHeight(42); self.finish_btn.clicked.connect(self.finish_to_layer); rl.addWidget(self.finish_btn)
        split.addWidget(right); split.setStretchFactor(0,4); split.setStretchFactor(1,1); split.setSizes([790,270])
        root.addWidget(split,1)
        self.setStyleSheet("""
        QDialog { background:#f3f6fb; color:#182230; font-family:'Segoe UI'; font-size:10pt; }
        QLabel { color:#243247; }
        QLabel#metaTitle { color:#102a43; font-size:17pt; font-weight:800; }
        QFrame#metaPanel { background:#ffffff; border:1px solid #dbe4ef; border-radius:10px; }
        QFrame#commandBar { background:#ffffff; border:1px solid #dbe4ef; border-radius:10px; }
        QToolButton { padding:8px 14px; border:1px solid #cbd6e4; border-radius:8px; background:#ffffff; min-height:22px; color:#243247; font-weight:700; }
        QToolButton:hover { background:#edf5ff; border-color:#82b1e8; }
        QMenu { background:#ffffff; border:1px solid #cbd6e4; padding:5px; }
        QMenu::item { padding:8px 28px 8px 12px; border-radius:5px; }
        QMenu::item:selected { background:#e7f1ff; color:#135da8; }
        QFrame#leftPanel, QFrame#rightPanel { background:#ffffff; border:1px solid #dce5ef; border-radius:12px; }
        QLabel#targetCard { background:#edf5ff; border:1px solid #c8ddf7; border-radius:12px; padding:10px; color:#135da8; font-size:25pt; font-weight:800; min-height:56px; }
        QPushButton#nextNumber { background:#e7f8ee; border:1px solid #a5dfba; border-radius:10px; color:#147542; font-size:24pt; font-weight:800; }
        QPushButton#nextNumber:hover { background:#d7f3e2; border-color:#51b979; }
        QLabel#statusCard { background:#f7f9fc; border:1px solid #e1e7f0; border-radius:9px; padding:10px; }
        QTabWidget::pane { border:1px solid #d7e0eb; border-radius:9px; background:#ffffff; top:-1px; }
        QTabBar::tab { padding:10px 16px; margin-right:4px; border:1px solid #dce4ef; border-bottom:none; border-top-left-radius:8px; border-top-right-radius:8px; background:#e9eef6; color:#526176; min-width:95px; font-weight:600; }
        QTabBar::tab:selected { background:#ffffff; color:#1769aa; font-weight:800; }
        QPushButton { padding:8px 12px; border:1px solid #cbd6e4; border-radius:8px; background:#ffffff; min-height:22px; color:#243247; font-weight:600; }
        QPushButton:hover { background:#edf5ff; border-color:#82b1e8; }
        QPushButton:pressed { background:#dcecff; }
        QPushButton:default { background:#1769aa; color:#ffffff; border-color:#1769aa; }
        QComboBox, QLineEdit, QSpinBox, QDoubleSpinBox { background:#ffffff; border:1px solid #cbd6e4; border-radius:7px; padding:7px 9px; min-height:20px; selection-background-color:#c7e2ff; }
        QComboBox:focus, QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus { border:1px solid #3984d6; }
        QTableWidget, QTreeWidget { background:#ffffff; alternate-background-color:#f6f9fd; gridline-color:#e6ebf2; selection-background-color:#cfe5ff; selection-color:#10243b; border:1px solid #dce4ef; border-radius:7px; outline:0; }
        QHeaderView::section { background:#edf2f8; color:#42536a; padding:8px 7px; border:0; border-right:1px solid #dce4ef; border-bottom:1px solid #dce4ef; font-weight:700; }
        QProgressBar { border:1px solid #ccd7e5; border-radius:7px; background:#edf2f7; text-align:center; min-height:21px; color:#233247; font-weight:700; }
        QProgressBar::chunk { background:#20a464; border-radius:6px; }
        QSplitter::handle { background:#e1e8f1; width:5px; }
        QToolTip { background:#102a43; color:#ffffff; border:1px solid #294b68; padding:5px 7px; }
        """)
        self._build_manual(); self._build_auto()
    def _load_style_from_popup(self):
        # Keep existing style implementation as the single source of behavior.
        self.load_style_current()

    def start_note_pick(self):
        layer = self.iface.activeLayer()
        if layer is None or not isinstance(layer, QgsVectorLayer) or QgsWkbTypes.geometryType(layer.wkbType()) != QgsWkbTypes.PointGeometry:
            QMessageBox.warning(self, 'Catatan Titik Bangunan', 'Pilih layer titik SLS pada panel Layers QGIS terlebih dahulu.')
            return
        canvas = self.iface.mapCanvas()
        self._note_old_tool = canvas.mapTool()
        parent = self
        class NotePickTool(QgsMapToolEmitPoint):
            def canvasReleaseEvent(tool, event):
                try:
                    map_point = tool.toMapCoordinates(event.pos())
                    click_geom = QgsGeometry.fromPointXY(QgsPointXY(map_point))
                    # Search within a screen-space tolerance, converted to map units.
                    tol_canvas = canvas.mapUnitsPerPixel() * 10.0
                    search_geom = QgsGeometry(click_geom)
                    if canvas.mapSettings().destinationCrs() != layer.crs():
                        tr = QgsCoordinateTransform(canvas.mapSettings().destinationCrs(), layer.crs(), QgsProject.instance())
                        search_geom.transform(tr)
                        # Approximate tolerance in layer units from the transformed canvas extent.
                        extent = canvas.extent()
                        corners = [QgsPointXY(extent.xMinimum(), extent.yMinimum()), QgsPointXY(extent.xMaximum(), extent.yMaximum())]
                        tr2 = QgsCoordinateTransform(canvas.mapSettings().destinationCrs(), layer.crs(), QgsProject.instance())
                        corners = [tr2.transform(pt) for pt in corners]
                        tol_layer = max(abs(corners[1].x()-corners[0].x())/max(canvas.width(),1), abs(corners[1].y()-corners[0].y())/max(canvas.height(),1))*10.0
                    else:
                        tol_layer = tol_canvas
                    rect = search_geom.boundingBox()
                    rect.grow(max(tol_layer, 1e-9))
                    request = QgsFeatureRequest().setFilterRect(rect)
                    best = None; bestdist = float('inf')
                    for f in layer.getFeatures(request):
                        if not f.hasGeometry() or f.geometry().isEmpty():
                            continue
                        d = f.geometry().distance(search_geom)
                        if d < bestdist:
                            bestdist=d; best=f
                    if best is None or bestdist > max(tol_layer * 2.0, 1e-8):
                        QMessageBox.warning(parent, 'Catatan Titik Bangunan', 'Titik tidak ditemukan dalam jangkauan klik. Perbesar peta dan klik dekat titik yang dimaksud.')
                        return
                    layer.removeSelection(); layer.selectByIds([best.id()]); parent.iface.setActiveLayer(layer)
                    if parent._note_old_tool:
                        canvas.setMapTool(parent._note_old_tool)
                    else:
                        canvas.unsetMapTool(tool)
                    fields = [f.name() for f in layer.fields()]
                    status_field = next((x for x in ['STATUS_CATATAN','status_catatan','STATUS_CAT','status_cat'] if x in fields), None)
                    note_field = next((x for x in ['CATATAN_TITIK','catatan_titik','CATATAN','catatan'] if x in fields), None)
                    status, ok = QInputDialog.getItem(parent, 'Catatan Titik Bangunan', 'Status masalah:', ['Perlu verifikasi','Noktah tidak ditemukan','Nomor tidak terbaca','Noktah dan nomor tidak terdeteksi','Masalah lainnya'], 0, False)
                    if not ok:
                        return
                    note, ok = QInputDialog.getMultiLineText(parent, 'Catatan Titik Bangunan', 'Keterangan untuk FID {}:'.format(best.id()), '')
                    if not ok:
                        return
                    started = not layer.isEditable()
                    if started and not layer.startEditing():
                        QMessageBox.warning(parent, 'Catatan Titik Bangunan', 'Layer tidak dapat diedit. Periksa izin file atau sumber data.')
                        return
                    # Create the standard fields if missing, so the feature works on ordinary editable layers.
                    new_fields = []
                    if not status_field:
                        new_fields.append(QgsField('STATUS_CATATAN', QVariant.String, len=80)); status_field = 'STATUS_CATATAN'
                    if not note_field:
                        new_fields.append(QgsField('CATATAN_TITIK', QVariant.String, len=0)); note_field = 'CATATAN_TITIK'
                    if new_fields:
                        for new_field in new_fields:
                            if not layer.addAttribute(new_field):
                                if started: layer.rollBack()
                                QMessageBox.warning(parent, 'Kolom catatan gagal dibuat', 'Sumber layer tidak mengizinkan penambahan kolom. Gunakan layer yang dapat diedit atau tambahkan STATUS_CATATAN dan CATATAN_TITIK secara manual.')
                                return
                        layer.updateFields()
                    layer.changeAttributeValue(best.id(), layer.fields().indexFromName(status_field), status)
                    layer.changeAttributeValue(best.id(), layer.fields().indexFromName(note_field), note)
                    # Commit only edits started by this tool. If the layer was already in edit
                    # mode, preserve the user's existing edit session and clearly report pending edits.
                    if started and not layer.commitChanges():
                        errors = '; '.join(layer.commitErrors())
                        layer.rollBack()
                        QMessageBox.warning(parent, 'Catatan Titik Bangunan', 'Gagal menyimpan perubahan: ' + errors)
                        return
                    layer.updateFeature(best)
                    layer.triggerRepaint()
                    if started:
                        QMessageBox.information(parent, 'Catatan Titik Bangunan', 'Catatan berhasil disimpan ke sumber data untuk FID {}.'.format(best.id()))
                    else:
                        QMessageBox.information(parent, 'Catatan Titik Bangunan', 'Catatan untuk FID {} sudah masuk ke edit buffer layer. Karena layer sudah dalam mode edit, klik Simpan Editan pada QGIS untuk menyimpan ke sumber data.'.format(best.id()))
                except Exception as exc:
                    QMessageBox.warning(parent, 'Catatan Titik Bangunan', 'Gagal memilih/menyimpan titik: {}'.format(exc))
        self._note_pick_tool = NotePickTool(canvas)
        canvas.setMapTool(self._note_pick_tool)
        self.status_label.setText('Catatan Titik: klik titik bermasalah pada peta untuk mengisi status dan keterangan.')

    def _open_audit_from_dialog(self):
        # Open the existing audit dialog as a separate floating window.
        try:
            dlg = PointAuditDialog(self.iface, self)
            dlg.exec_()
        except Exception as exc:
            QMessageBox.warning(self, 'Audit Laporan Titik', 'Audit tidak dapat dibuka: {}'.format(exc))

    def open_georeference_converter(self):
        dlg = GeoreferenceConvertDialog(self.iface, self)
        dlg.exec_()

    def load_style_current(self):
        if not self.source:
            QMessageBox.information(self, 'Pilih layer', 'Pilih layer titik bangunan terlebih dahulu.'); return
        self.apply_style(self.source)
        self.status_label.setText('Style diterapkan: satu warna seragam, ukuran titik 4, label no_bang putih tebal dengan buffer hitam.')

    def _build_manual(self):
        l = self.manual_tab.layout()
        hint=QLabel('ALUR MANUAL  •  Klik baris nomor → klik noktah hitam pada peta. Setelah satu titik, daftar maju otomatis ke nomor berikutnya. Klik baris lain kapan saja untuk mengoreksi urutan.')
        hint.setWordWrap(True); l.addWidget(hint)
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(['FID', 'no_bang', 'Kode desa', 'Nama SLS', 'flag', 'Status', 'Koordinat target'])
        self._setup_table(self.table)
        l.addWidget(self.table, 1)
        controls = QHBoxLayout()
        self.next_btn = QPushButton('Pilih nomor belum dipetakan'); self.next_btn.clicked.connect(self.select_next_unmapped)
        controls.addWidget(self.next_btn)
        self.clear_btn = QPushButton('Hapus target baris terpilih'); self.clear_btn.clicked.connect(self.clear_selected)
        controls.addWidget(self.clear_btn)
        l.addLayout(controls)
        self.table.itemSelectionChanged.connect(self.arm_selected_row)

    def _build_auto(self):
        l = self.auto_tab.layout()
        l.addWidget(QLabel('Algoritma mencocokkan atribut no_bang dengan nomor satu digit pada raster Georeference (contoh no_bang 001/0001 cocok dengan label raster 1). OCR digunakan untuk membaca angka; transparansi tampilan GeoTIFF di QGIS (mis. 20%) tidak mengubah piksel sumber jika raster dibaca langsung; kandidat wajib ditinjau sebelum disimpan.'))
        settings = QHBoxLayout()
        settings.addWidget(QLabel('Radius pencarian (meter):'))
        self.radius = QDoubleSpinBox(); self.radius.setRange(0.5, 30.0); self.radius.setValue(5.0); self.radius.setSingleStep(0.5); self.radius.setSuffix(' m')
        settings.addWidget(self.radius)
        settings.addWidget(QLabel('Maks. sampel per sisi:'))
        self.samples = QSpinBox(); self.samples.setRange(5, 99); self.samples.setValue(40); self.samples.setToolTip('Nilai 5–99; ukuran citra OCR internal dibatasi untuk menjaga kinerja.')
        settings.addWidget(self.samples)
        settings.addWidget(QLabel('Raster referensi:'))
        self.raster_combo = QComboBox(); self.raster_combo.setMinimumWidth(210); self.raster_combo.setMaxVisibleItems(30); self.raster_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed); self.raster_combo.setToolTip('Pilih raster lokal dari grup Georeference; pilihan mengikuti Raster SLS jika tersedia.'); settings.addWidget(self.raster_combo, 1)
        self.raster_combo.currentIndexChanged.connect(self._on_auto_raster_selected)
        self.raster_choose_btn = QPushButton('Choose')
        self.raster_choose_btn.setToolTip('Pilih raster dari grup Georeference dan aktifkan layer tersebut di panel QGIS')
        self.raster_choose_btn.clicked.connect(lambda: self._choose_georeference_raster('auto'))
        settings.addWidget(self.raster_choose_btn)
        self.auto_btn = QPushButton('🔎 Cari saran otomatis'); self.auto_btn.setToolTip('Baca nomor pada raster dengan OCR lalu cocokkan dengan no_bang; 001/0001 cocok dengan label 1'); self.auto_btn.clicked.connect(self.run_auto_suggestions)
        settings.addWidget(self.auto_btn)
        l.addLayout(settings)
        self.auto_table = QTableWidget(0, 7)
        self.auto_table.setHorizontalHeaderLabels(['FID', 'no_bang', 'Kode desa', 'Nama SLS', 'flag', 'Saran otomatis', 'Skor OCR / jarak'])
        self._setup_table(self.auto_table)
        self.auto_table.itemSelectionChanged.connect(self._zoom_selected_auto_flag)
        l.addWidget(self.auto_table, 1)
        buttons = QHBoxLayout()
        self.apply_auto_btn = QPushButton('Terapkan saran terpilih'); self.apply_auto_btn.clicked.connect(self.apply_selected_suggestions)
        buttons.addWidget(self.apply_auto_btn)
        self.correct_btn = QPushButton('Koreksi dengan klik peta'); self.correct_btn.clicked.connect(self.start_pick_from_auto)
        buttons.addWidget(self.correct_btn)
        l.addLayout(buttons)

    def _setup_table(self, table):
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        table.setWordWrap(False)
        for col in range(table.columnCount()-1): table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(table.columnCount()-1, QHeaderView.Stretch)

    def _on_primary_group_changed(self, *_):
        """Village code is the parent selector; require a fresh SLS choice per village."""
        self.group_sls_combo.blockSignals(True)
        if self.group_sls_combo.count():
            self.group_sls_combo.setCurrentIndex(0)
        self.group_sls_combo.blockSignals(False)
        self._refresh_sls_filter()
        self._apply_group_filter()

    def _refresh_sls_filter(self):
        if not hasattr(self, 'group_sls_combo'):
            return
        village_code = self.group_primary_combo.currentData()
        previous = self.group_sls_combo.currentData()
        names = sorted({
            v.get('sls', '') for v in self._feature_group_values.values()
            if v.get('sls') and village_code is not None and v.get('desa_code') == village_code
        }, key=lambda x: x.casefold())
        self.group_sls_combo.blockSignals(True)
        self.group_sls_combo.clear()
        self.group_sls_combo.addItem('Pilih SLS…', None)
        for name in names:
            self.group_sls_combo.addItem(name, name)
        idx = self.group_sls_combo.findData(previous)
        if idx > 0:
            self.group_sls_combo.setCurrentIndex(idx)
        self.group_sls_combo.setEnabled(village_code is not None and bool(names))
        self.group_sls_combo.setVisible(True)
        self.group_sls_label.setVisible(True)
        self.group_sls_combo.blockSignals(False)

    def _base_subset_for(self, layer):
        if not layer:
            return ''
        lid = layer.id()
        if lid not in self._original_subset_strings:
            try:
                self._original_subset_strings[lid] = layer.subsetString()
            except Exception:
                self._original_subset_strings[lid] = ''
        return self._original_subset_strings[lid]

    def _restore_layer_subset(self, layer):
        if not layer:
            return
        base = self._original_subset_strings.get(layer.id())
        if base is None:
            return
        try:
            provider = layer.dataProvider()
            current = layer.subsetString()
            if current != base:
                provider.setSubsetString(base)
                layer.updateExtents()
                layer.triggerRepaint()
        except Exception:
            pass

    def restore_source_subset(self):
        """Restore original provider filter when the AutoSnap window is closed."""
        if self.source:
            self._restore_layer_subset(self.source)
            self.iface.mapCanvas().refresh()
        self._subset_filter_applied = False

    @staticmethod
    def _sql_quote(value):
        return "'" + str(value).replace("'", "''") + "'"

    def _apply_map_subset(self, visible_fids):
        """Temporarily show only the chosen SLS in the QGIS map; never edits source features."""
        if not self.source:
            return
        layer = self.source
        base = self._base_subset_for(layer)
        try:
            provider = layer.dataProvider()
            if not provider.supportsSubsetString():
                self.status_label.setText('Filter peta tidak didukung provider layer ini. Daftar tools tetap difilter, tetapi layer sumber di peta tidak dapat disembunyikan per fitur oleh provider tersebut.')
                return
            village_code = self.group_primary_combo.currentData()
            sls_name = self.group_sls_combo.currentData()
            if village_code is None or sls_name is None or not visible_fids:
                condition = '1=0'
            else:
                # Use the exact original full codes belonging to the selected village/SLS.
                full_codes = sorted({
                    v.get('full_code', '') for fid, v in self._feature_group_values.items()
                    if int(fid) in visible_fids and v.get('full_code')
                })
                names = [f.name() for f in layer.fields()]
                code_field = names[self._field_index(names, ('level_6_full_code',))] if self._field_index(names, ('level_6_full_code',)) >= 0 else ''
                sls_idx = self._field_index(names, ('level_6_name','nama_sls','sls_name','nama_satuan_lingkungan_setempat','sls'))
                sls_field = names[sls_idx] if sls_idx >= 0 else ''
                if not code_field or not sls_field or not full_codes:
                    # FID fallback for providers with the expected stable fid column.
                    ids = ','.join(str(int(fid)) for fid in sorted(visible_fids))
                    condition = f'"fid" IN ({ids})' if ids else '1=0'
                else:
                    values = ','.join(self._sql_quote(v) for v in full_codes)
                    condition = f'"{code_field}" IN ({values}) AND "{sls_field}" = {self._sql_quote(sls_name)}'
            expression = f'({base}) AND ({condition})' if base.strip() else condition
            current = layer.subsetString()
            if current != expression:
                ok = provider.setSubsetString(expression)
                if ok is False:
                    self.status_label.setText('Filter peta gagal diterapkan oleh provider. Daftar tools tetap difilter; periksa dukungan filter layer sumber.')
                    return
                layer.updateExtents()
                layer.triggerRepaint()
            self._subset_filter_applied = True
            self.iface.mapCanvas().refresh()
            if village_code is not None and sls_name is not None and visible_fids:
                try:
                    ext = layer.extent()
                    if not ext.isNull() and not ext.isEmpty():
                        ext.scale(1.25)
                        self.iface.mapCanvas().setExtent(ext)
                        self.iface.mapCanvas().refresh()
                except Exception:
                    pass
        except Exception as e:
            self.status_label.setText(f'Filter peta tidak dapat diterapkan: {type(e).__name__}: {e}. Daftar tools tetap difilter.')

    def _is_georeference_raster(self, layer, node=None):
        """Recognize georeference rasters by their layer name or containing tree group."""
        if not isinstance(layer, QgsRasterLayer):
            return False
        name = (layer.name() or '').casefold()
        if any(token in name for token in ('georeference', 'georeferenced', '_wss', ' wss')):
            return True
        try:
            node = node or QgsProject.instance().layerTreeRoot().findLayer(layer.id())
            parent = node.parent() if node else None
            while parent is not None:
                if 'georeference' in (parent.name() or '').casefold():
                    return True
                parent = parent.parent()
        except Exception:
            pass
        return False

    @staticmethod
    def _normalize_raster_sls_name(value):
        """Normalize a raster/layer name for exact SLS-name matching."""
        value = os.path.basename(str(value or '').strip())
        value = re.sub(r'\.(tif|tiff|png|jpg|jpeg|vrt|img|jp2|asc)$', '', value, flags=re.I)
        value = re.sub(r'(?:[_\s-]+(?:WSS|GEOREFERENCED|GEOREFERENCE))+$', '', value, flags=re.I)
        return re.sub(r'\s+', ' ', value).strip().casefold()

    def _rasters_in_georeference_group(self):
        """Return valid layers inside any Georeference group.

        Some project files load the imagery/reference entries as non-raster providers,
        so do not reject them solely because the provider is not QgsRasterLayer.
        Prefer true raster layers, then include other valid descendants as a fallback.
        """
        project = QgsProject.instance()
        root = project.layerTreeRoot()
        rasters = []
        fallback = []
        for layer in project.mapLayers().values():
            if not layer or not layer.isValid():
                continue
            node = root.findLayer(layer.id())
            parent = node.parent() if node else None
            in_group = False
            while parent is not None and parent != root:
                if hasattr(parent, 'name') and 'georeference' in str(parent.name() or '').strip().casefold():
                    in_group = True
                    break
                parent = parent.parent()
            if in_group:
                fallback.append(layer)
                if isinstance(layer, QgsRasterLayer):
                    rasters.append(layer)
        return rasters if rasters else fallback

    def _georeference_group_label_for_layer(self, layer):
        """Return the nearest named group inside the Georeference tree for a layer."""
        try:
            root = QgsProject.instance().layerTreeRoot()
            node = root.findLayer(layer.id())
            parent = node.parent() if node else None
            georef_found = False
            group_labels = []
            while parent is not None and parent != root:
                name = str(parent.name() or '').strip()
                if 'georeference' in name.casefold():
                    georef_found = True
                    break
                group_labels.append(name)
                parent = parent.parent()
            if not georef_found:
                return layer.name()
            # Prefer the immediate group name above the raster, e.g. the SLS group.
            return group_labels[0] if group_labels and group_labels[0] else layer.name()
        except Exception:
            return layer.name()

    def _populate_georeference_raster_combos(self):
        """Populate selectors with SLS group names; each option still points to its raster layer."""
        rasters = self._rasters_in_georeference_group()
        for combo in (getattr(self, 'raster_sls_combo', None), getattr(self, 'raster_combo', None)):
            if combo is None:
                continue
            old_id = combo.currentData()
            combo.blockSignals(True)
            combo.clear()
            combo.addItem('Pilih grup SLS Georeference…', None)
            used_labels = {}
            for layer in sorted(rasters, key=lambda lyr: (self._georeference_group_label_for_layer(lyr).casefold(), lyr.name().casefold())):
                group_label = self._georeference_group_label_for_layer(layer)
                # If a group has multiple raster children, distinguish them without losing
                # the group name used for SLS matching.
                count = used_labels.get(group_label, 0)
                used_labels[group_label] = count + 1
                label = group_label if count == 0 else f'{group_label} — {layer.name()}'
                combo.addItem(label, layer.id())
            idx = combo.findData(old_id) if old_id else -1
            if idx < 0 and rasters:
                visible = next((r for r in rasters if self._layer_is_visible(r)), None)
                idx = combo.findData(visible.id()) if visible else 1
            combo.setCurrentIndex(idx if idx >= 0 else 0)
            combo.setEnabled(True)
            combo.setToolTip('Pilihan menampilkan nama grup SLS di dalam grup Georeference; layer raster terkait akan diaktifkan.')
            combo.blockSignals(False)

    def _choose_georeference_raster(self, target='sls'):
        """Explicit Choose action: select a raster from Georeference and activate it in QGIS."""
        combo = self.raster_sls_combo if target == 'sls' else getattr(self, 'raster_combo', None)
        rasters = self._rasters_in_georeference_group()
        if not rasters:
            QMessageBox.warning(self, 'Georeference kosong',
                'Tidak ditemukan layer raster valid di dalam grup Georeference. Periksa nama grup dan pastikan layer berupa raster.')
            return
        names = [r.name() for r in rasters]
        current = combo.currentText() if combo else ''
        initial = names.index(current) if current in names else 0
        chosen_name, ok = QInputDialog.getItem(self, 'Pilih Raster Georeference',
            'Raster dari grup Georeference:', names, initial, False)
        if not ok or not chosen_name:
            return
        layer = next((r for r in rasters if r.name() == chosen_name), None)
        if not layer:
            return
        for c in (getattr(self, 'raster_sls_combo', None), getattr(self, 'raster_combo', None)):
            if c is not None:
                idx = c.findData(layer.id())
                if idx >= 0:
                    c.blockSignals(True); c.setCurrentIndex(idx); c.blockSignals(False)
        self._set_georeference_raster_visible(layer.id())
        try:
            self.iface.setActiveLayer(layer)
            self.iface.layerTreeView().setCurrentLayer(layer)
        except Exception:
            pass
        try:
            self.status_label.setText(f'Raster aktif: {layer.name()}')
        except Exception:
            pass

    def _on_auto_raster_selected(self, *_):
        """Use only rasters inside the Georeference layer-tree group."""
        if not getattr(self, 'raster_combo', None) or self.raster_combo.signalsBlocked():
            return
        layer_id = self.raster_combo.currentData()
        if not layer_id:
            return
        project = QgsProject.instance()
        layer = project.mapLayer(layer_id)
        if layer not in self._rasters_in_georeference_group():
            return
        if hasattr(self, 'raster_sls_combo'):
            idx = self.raster_sls_combo.findData(layer_id)
            if idx >= 0:
                self.raster_sls_combo.blockSignals(True)
                self.raster_sls_combo.setCurrentIndex(idx)
                self.raster_sls_combo.blockSignals(False)
        self._set_georeference_raster_visible(layer_id)

    def _set_georeference_raster_visible(self, layer_id):
        project = QgsProject.instance()
        root = project.layerTreeRoot()
        allowed = {r.id() for r in self._rasters_in_georeference_group()}
        if layer_id not in allowed:
            return
        for rid in allowed:
            node = root.findLayer(rid)
            if node is not None:
                if rid not in self._georef_original_visibility:
                    self._georef_original_visibility[rid] = bool(node.isVisible())
                self._georef_managed_ids.add(rid)
                node.setItemVisibilityChecked(rid == layer_id)
        node = root.findLayer(layer_id)
        parent = node.parent() if node else None
        while parent is not None and parent != root:
            if hasattr(parent, 'setItemVisibilityChecked'):
                parent.setItemVisibilityChecked(True)
            parent = parent.parent()
        self.iface.mapCanvas().refresh()

    def _on_raster_sls_selected(self, *_):
        """Manual raster selection; only the active Georeference group is eligible."""
        if not hasattr(self, 'raster_sls_combo') or self.raster_sls_combo.signalsBlocked():
            return
        layer_id = self.raster_sls_combo.currentData()
        if not layer_id:
            return
        project = QgsProject.instance()
        selected = project.mapLayer(layer_id)
        if selected not in self._rasters_in_georeference_group():
            return
        self._set_georeference_raster_visible(layer_id)
        if getattr(self, 'raster_combo', None):
            idx = self.raster_combo.findData(layer_id)
            if idx >= 0:
                self.raster_combo.blockSignals(True)
                self.raster_combo.setCurrentIndex(idx)
                self.raster_combo.blockSignals(False)
        try:
            self.status_label.setText(self.status_label.text() + f' Raster dipilih: {selected.name()}.')
        except Exception:
            pass

    def _sync_georeference_visibility(self, village_code, sls_name):
        """Refresh raster choices without overriding the user's explicit raster selection.

        The user chooses a raster from the Georeference layers (the layers indicated
        in the QGIS Layers panel). The selected raster is checked and the other
        rasters in Georeference groups are unchecked by the selector handlers.
        """
        try:
            selected_ids = []
            for combo in (getattr(self, 'raster_sls_combo', None),
                          getattr(self, 'raster_combo', None)):
                if combo is not None and combo.currentData():
                    selected_ids.append(combo.currentData())

            self._populate_georeference_raster_combos()
            rasters = self._rasters_in_georeference_group()
            if not rasters:
                return

            # Match the selected SLS against the actual group name in the QGIS
            # Layers tree first; use the raster layer name only as a fallback.
            wanted = self._normalize_raster_sls_name(sls_name)
            chosen_id = None
            if wanted:
                for raster in rasters:
                    group_name = self._normalize_raster_sls_name(self._georeference_group_label_for_layer(raster))
                    if group_name == wanted:
                        chosen_id = raster.id()
                        break
                if chosen_id is None:
                    for raster in rasters:
                        if self._normalize_raster_sls_name(raster.name()) == wanted:
                            chosen_id = raster.id()
                            break
            if chosen_id is None:
                chosen_id = next((rid for rid in selected_ids
                                  if any(r.id() == rid for r in rasters)), None)
            if chosen_id:
                for combo in (getattr(self, 'raster_sls_combo', None),
                              getattr(self, 'raster_combo', None)):
                    if combo is not None:
                        idx = combo.findData(chosen_id)
                        if idx >= 0:
                            combo.blockSignals(True)
                            combo.setCurrentIndex(idx)
                            combo.blockSignals(False)
                self._set_georeference_raster_visible(chosen_id)
        except Exception as e:
            try:
                self.status_label.setText(
                    self.status_label.text() + f' Raster Georeference: {e}')
            except Exception:
                pass

    def restore_georeference_visibility(self):
        """Restore the user's original raster visibility when the dialog closes."""
        try:
            root = QgsProject.instance().layerTreeRoot()
            for layer_id, visible in self._georef_original_visibility.items():
                node = root.findLayer(layer_id)
                if node is not None:
                    node.setItemVisibilityChecked(bool(visible))
            self.iface.mapCanvas().refresh()
        except Exception:
            pass

    def _apply_group_filter(self, *_):
        """Filter tables by village code + SLS name and filter the map to that same SLS."""
        if not hasattr(self, 'group_primary_combo') or not self._feature_group_values or self._building_tables:
            return
        village_code = self.group_primary_combo.currentData()
        selected_sls = self.group_sls_combo.currentData() if self.group_sls_combo.isEnabled() else None
        visible_fids = []
        for fid, values in self._feature_group_values.items():
            if village_code is None or selected_sls is None:
                continue
            if values.get('desa_code') != village_code:
                continue
            if values.get('sls', '') != selected_sls:
                continue
            visible_fids.append(int(fid))
        self._selected_village_code = village_code
        self._selected_sls_name = selected_sls
        self._active_visible_fids = set(visible_fids)
        if hasattr(self, 'total_points_btn'):
            self.total_points_btn.setText(f'TOTAL TITIK SLS AKTIF  •  {len(visible_fids):,}')
            self.total_points_btn.setEnabled(village_code is not None and selected_sls is not None)
        if village_code is not None:
            village_values = [v for v in self._feature_group_values.values() if v.get('desa_code') == village_code]
            village_name = next((v.get('desa') for v in village_values if v.get('desa')), village_code)
            self.desa_edit.setText(str(village_name))
        else:
            self.desa_edit.setText('')
        self.sls_edit.setText(str(selected_sls or ''))
        self._populate_group_tables(visible_fids)
        self._apply_map_subset(set(visible_fids))
        if village_code is None:
            self.status_label.setText('Pilih kode/nama desa terlebih dahulu, lalu pilih SLS. Fitur lain disembunyikan dari peta selama workflow aktif.')
        elif selected_sls is None:
            self.status_label.setText('Desa dipilih. Sekarang pilih nama SLS di sebelah kanan untuk menampilkan hanya titik pada SLS tersebut.')
        else:
            village_values = [v for v in self._feature_group_values.values() if v.get('desa_code') == village_code]
            village_name = next((v.get('desa') for v in village_values if v.get('desa')), village_code)
            self.status_label.setText(f'Desa: {village_name} ({village_code}) • SLS: {selected_sls}. {len(visible_fids):,} titik aktif; titik SLS lain disembunyikan di peta.')
        self._sync_georeference_visibility(village_code, selected_sls)
        self._update_target_card(self._active_fid)

    def show_active_group_count(self):
        count = len(getattr(self, '_active_visible_fids', set()))
        code = self.group_primary_combo.currentData() if hasattr(self, 'group_primary_combo') else None
        sls = self.group_sls_combo.currentData() if hasattr(self, 'group_sls_combo') else None
        if not code or not sls:
            QMessageBox.information(self, 'Total titik', 'Pilih kode desa dan SLS terlebih dahulu.')
            return
        QMessageBox.information(self, 'Total titik pada SLS aktif',
            f'Kode desa: {code}\nNama SLS: {sls}\nTotal titik: {count:,}')

    def _populate_group_tables(self, visible_fids):
        """Populate visible-group rows only. Source attributes are cached without geometry."""
        self._building_tables = True
        try:
            records = [self._feature_records[fid] for fid in visible_fids if fid in self._feature_records]
            records.sort(key=lambda rec: (str(rec['no']).strip().zfill(8), rec['fid']))
            self._row_lookup = {self.table: {}, self.auto_table: {}}
            for table in (self.table, self.auto_table):
                table.setUpdatesEnabled(False)
                table.blockSignals(True)
                table.setRowCount(0)
                table.setRowCount(len(records))
            for row, rec in enumerate(records):
                fid, no = rec['fid'], rec['no']
                mapped = fid in self.mapping
                target = self.mapping.get(fid)
                status = 'Target manual' if mapped else 'Belum dipetakan'
                coord = f'{target.x():.8f}, {target.y():.8f}' if target else '—'
                auto = self._auto_candidates.get(fid)
                auto_status = 'Saran tersedia' if auto else ('Belum dianalisis' if not mapped else 'Target sudah dipetakan')
                auto_coord = f'{auto[0].x():.8f}, {auto[0].y():.8f}' if auto else '—'
                # The workflow flag is 1 until this feature receives a target; then it becomes 0.
                flag_value = '0' if mapped else str(rec.get('source_flag') or '1')
                common = [str(fid), no, str(rec.get('desa_code') or '—'), str(rec.get('sls') or '—'), flag_value]
                for table, values in ((self.table, common + [status, coord]),
                                      (self.auto_table, common + [auto_status, auto_coord])):
                    self._row_lookup[table][fid] = row
                    for col, value in enumerate(values):
                        item = QTableWidgetItem(str(value))
                        if col == 0:
                            item.setData(Qt.UserRole, fid)
                        if col == 1 and rec.get('duplicate', False):
                            item.setBackground(QColor('#fff0b3'))
                            item.setToolTip('Nomor bangunan duplikat; kerjakan tiap fitur/FID secara terpisah.')
                        if col == 4:
                            item.setTextAlignment(Qt.AlignCenter)
                            item.setForeground(QColor('#15803d' if flag_value == '0' else '#b45309'))
                        table.setItem(row, col, item)
            for table in (self.table, self.auto_table):
                table.setAlternatingRowColors(True)
                table.blockSignals(False)
                table.setUpdatesEnabled(True)
                table.viewport().update()
        finally:
            self._building_tables = False
        mapped_in_group = sum(1 for fid in visible_fids if fid in self.mapping)
        pct = int(mapped_in_group * 100 / max(1, len(visible_fids)))
        self.process_progress.setValue(pct)
        self.process_progress.setFormat(f'PEMETAAN • {mapped_in_group}/{len(visible_fids)}')
        self.circular_progress.setValue(pct)

    def _resolve_village(self, full_code):
        digits=''.join(re.findall(r'\d', str(full_code or '')))
        # level_6_full_code in this dataset is 16 digits: district prefix + local
        # village sequence + SLS/block suffix. The PDF uses the 10-digit village
        # code, where the village suffix starts with 2 (e.g. 3210202001).
        candidates=[]
        if len(digits) >= 10:
            if len(digits) >= 10 and digits[6] == '0':
                candidates.append(digits[:6] + '2' + digits[7:10])
            candidates.append(digits[:10])
            for i in range(max(0,len(digits)-9)):
                candidates.append(digits[i:i+10])
        for code in candidates:
            if code in VILLAGE_BY_CODE:
                name, kec_code, kec_name = VILLAGE_BY_CODE[code]
                return code, name, kec_code, kec_name
        fallback = digits[:10] if len(digits)>=10 else digits or 'KODE TIDAK TERSEDIA'
        return fallback, 'DESA TIDAK DITEMUKAN ('+fallback+')', digits[:6], ''

    def _setup_group_filters(self, feats, names):
        """Create village-code selector from level_6_full_code and dependent SLS selector."""
        self._feature_group_values = {}
        self._feature_records = {}
        code_idx = self._field_index(names, ('level_6_full_code',))
        sls_idx = self._field_index(names, ('level_6_name','nama_sls','sls_name','nama_satuan_lingkungan_setempat','sls'))
        desa_idx = self._field_index(names, ('nama_desa','desa','desa_name','village','nama_kelurahan','kelurahan','nama_desa_kelurahan'))
        no_idx = self._field_index(names, ('no_bang',))
        counts = {}
        cached = []
        for feat in feats:
            fid = int(feat.id())
            full_code = str(feat[code_idx]).strip() if code_idx >= 0 and feat[code_idx] not in (None, '') else ''
            derived_code, derived_name, kec_code, kec_name = self._resolve_village(full_code)
            desa = str(feat[desa_idx]).strip() if desa_idx >= 0 and feat[desa_idx] not in (None, '') else derived_name
            sls = str(feat[sls_idx]).strip() if sls_idx >= 0 and feat[sls_idx] not in (None, '') else ''
            no = str(feat[no_idx]).strip() if no_idx >= 0 and feat[no_idx] not in (None, '') else ''
            self._feature_group_values[fid] = {
                'desa': desa, 'desa_code': derived_code, 'kecamatan_code': kec_code,
                'kecamatan': kec_name, 'sls': sls, 'full_code': full_code
            }
            flag_idx = self._field_index(names, ('flag',))
            raw_flag = str(feat[flag_idx]).strip() if flag_idx >= 0 and feat[flag_idx] not in (None, '') else '1'
            self._feature_records[fid] = {
                'fid': fid, 'no': no, 'duplicate': False,
                'desa_code': derived_code, 'sls': sls,
                'source_flag': raw_flag
            }
            counts[no] = counts.get(no, 0) + 1
            cached.append((fid, no))
        for fid, no in cached:
            self._feature_records[fid]['duplicate'] = counts.get(no, 0) > 1

        previous_code = self.group_primary_combo.currentData()
        self.group_primary_combo.blockSignals(True)
        self.group_primary_combo.clear()
        self.group_primary_combo.addItem('Pilih desa / kode desa…', None)
        villages = {}
        for values in self._feature_group_values.values():
            code = values.get('desa_code') or values.get('full_code') or ''
            if not code:
                continue
            entry = villages.setdefault(code, {'name': '', 'full_codes': set()})
            name = str(values.get('desa') or '').strip()
            if name and not name.startswith('DESA TIDAK DITEMUKAN') and not entry['name']:
                entry['name'] = name
            if values.get('full_code'):
                entry['full_codes'].add(values['full_code'])
        self._village_options = villages
        for code in sorted(villages):
            name = villages[code]['name']
            label = str(code)
            self.group_primary_combo.addItem(label, code)
        idx = self.group_primary_combo.findData(previous_code)
        if idx > 0:
            self.group_primary_combo.setCurrentIndex(idx)
        else:
            self.group_primary_combo.setCurrentIndex(0)
        self.group_primary_label.setText('🏘 Kode desa')
        self.group_sls_label.setText('▦ SLS')
        self.group_sls_combo.setVisible(True)
        self.group_sls_label.setVisible(True)
        self.group_primary_combo.blockSignals(False)
        self._refresh_sls_filter()
        self._apply_group_filter()

    def _layer_is_visible(self, layer):
        """Read visibility from the QGIS layer tree node (layers do not expose isVisible())."""
        try:
            node = QgsProject.instance().layerTreeRoot().findLayer(layer.id())
            return bool(node and node.isVisible())
        except Exception:
            return False

    def _is_web_raster(self, layer):
        if not layer: return True
        try:
            source=(layer.source() or '').lower(); name=layer.name().lower(); provider=(layer.providerType() or '').lower()
            return (provider in ('wms','wcs','arcgismapserver') or any(t in source for t in ('type=xyz','type=wmts','url=','tilematrixset','google.com/maps','googleapis.com','openstreetmap')) or any(t in name for t in ('google hybrid','google satellite','openstreetmap','xyz tiles','basemap','quickmapservices')))
        except Exception: return False

    def refresh_layers(self, force=False):
        """Refresh layer selectors cheaply; skip full data rebuild when project layer set is unchanged."""
        project_layers=list(QgsProject.instance().mapLayers().values())
        signature=tuple(sorted((l.id(), l.name(), l.isValid()) for l in project_layers))
        if not force and self._layer_signature == signature and self.source and self.source.isValid():
            return
        self._layer_signature=signature
        current = self.layer_combo.currentData()
        self.layer_combo.blockSignals(True); self.layer_combo.clear()
        layers=[]
        for lyr in project_layers:
            if isinstance(lyr,QgsVectorLayer) and lyr.isValid() and lyr.geometryType()==QgsWkbTypes.PointGeometry and lyr.fields().indexFromName('no_bang')>=0:
                layers.append(lyr)
        active=self.iface.activeLayer()
        layers.sort(key=lambda x:(x.id()!=(active.id() if active else ''),x.name().lower()))
        for lyr in layers:self.layer_combo.addItem(lyr.name(),lyr.id())
        ix=self.layer_combo.findData(current) if current else -1
        if ix<0 and active:
            ix=self.layer_combo.findData(active.id())
        if ix>=0:self.layer_combo.setCurrentIndex(ix)
        self.layer_combo.blockSignals(False)
        self._populate_georeference_raster_combos()
        selected_id=self.layer_combo.currentData()
        if not self.source or not self.source.isValid() or self.source.id()!=selected_id:
            self.load_features()

    def current_layer(self):
        lid = self.layer_combo.currentData()
        return QgsProject.instance().mapLayer(lid) if lid else None

    def load_features(self, *_):
        self.stop_pick(restore=False)
        previous_source = self.source
        selected_layer = self.current_layer()
        if previous_source and selected_layer and previous_source.id() != selected_layer.id():
            self._restore_layer_subset(previous_source)
        elif previous_source and selected_layer and previous_source.id() == selected_layer.id() and self._subset_filter_applied:
            # Rebuilding selectors must start from the complete source, not the last SLS subset.
            self._restore_layer_subset(previous_source)
            self._subset_filter_applied = False
        self.source=selected_layer; self.mapping={}; self._auto_candidates={}
        self._feature_records={}; self._row_lookup={}
        self.table.setRowCount(0); self.auto_table.setRowCount(0)
        if not self.source:
            self.status_label.setText('Tidak ditemukan layer titik dengan atribut no_bang.'); return
        self._base_subset_for(self.source)
        names=[fld.name() for fld in self.source.fields()]
        # Request only fields needed for the selector and tables, without geometries.
        needed=set()
        for candidates in (('no_bang',),('level_6_full_code',),('level_6_name','nama_sls','sls_name','nama_satuan_lingkungan_setempat','sls'),('nama_desa','desa','desa_name','village','nama_kelurahan','kelurahan','nama_desa_kelurahan'),('flag',)):
            idx=self._field_index(names,candidates)
            if idx>=0:needed.add(idx)
        request=QgsFeatureRequest().setFlags(QgsFeatureRequest.NoGeometry)
        if needed: request.setSubsetOfAttributes(sorted(needed))
        feats=list(self.source.getFeatures(request))
        if not feats:
            self.status_label.setText('Layer tidak memiliki fitur untuk ditampilkan.'); return
        self._setup_group_filters(feats,names)
        # Only the chosen village/SLS is rendered into the tables; never allocate 30k+ rows twice at startup.
        self._apply_group_filter()
        self.auto_progress.setValue(0)
        self.process_progress.setValue(0)
        self.status_label.setText(f'{len(feats):,} fitur terindeks dari {self.source.name()}. Pilih Desa dan SLS untuk memusatkan pekerjaan.')

    @staticmethod
    def _field_index(names, candidates):
        lookup = {str(n).strip().lower(): i for i, n in enumerate(names)}
        for candidate in candidates:
            if candidate.lower() in lookup: return lookup[candidate.lower()]
        # fallback partial matches for local schema variants
        for i, name in enumerate(names):
            low = str(name).strip().lower()
            if any(c.lower() in low for c in candidates): return i
        return -1

    def _update_target_card(self, fid=None):
        if not self.source or not hasattr(self, 'current_target_label'): return
        if fid is None:
            fid=self._active_fid
        if fid is not None:
            row=self._row_for_fid(self.table, fid)
            no=self.table.item(row,1).text() if row>=0 else '—'
            self.current_target_label.setText(no)
            self.current_target_label.setToolTip(f'TARGET AKTIF • FID {fid}. Klik noktah hitam kecil pada raster.')
        else:
            self.current_target_label.setText('—')
            self.current_target_label.setToolTip('Pilih baris fitur di sebelah kiri')
        nxt=None
        for r in range(self.table.rowCount()):
            f=int(self.table.item(r,0).text())
            if f not in self.mapping and f != fid:
                nxt=(f, self.table.item(r,1).text()); break
        if nxt:
            self.next_target_btn.setText(str(nxt[1]))
            self.next_target_btn.setEnabled(True)
            self.next_target_btn.setProperty('next_fid', nxt[0])
        else:
            self.next_target_btn.setText('✓')
            self.next_target_btn.setEnabled(False)

    def go_next_target(self):
        fid=self.next_target_btn.property('next_fid')
        if fid is None:
            self.select_next_unmapped(); return
        row=self._row_for_fid(self.table, int(fid))
        if row>=0:
            self.tabs.setCurrentWidget(self.manual_tab)
            self.table.selectRow(row); self.table.scrollToItem(self.table.item(row,0))
            self.arm_selected_row()

    def _row_for_fid(self, table, fid):
        try:
            row=self._row_lookup.get(table, {}).get(int(fid), -1)
            if row>=0 and row<table.rowCount(): return row
        except Exception: pass
        for r in range(table.rowCount()):
            item=table.item(r,0)
            if item and int(item.text())==int(fid): return r
        return -1

    def selected_fid(self, table=None):
        table = table or self.table; row = table.currentRow()
        return int(table.item(row,0).text()) if row >= 0 else None

    def start_pick(self):
        # Compatibility action: selecting a row arms the map click directly.
        self.arm_selected_row()

    def start_pick_from_auto(self):
        fid=self.selected_fid(self.auto_table)
        if fid is None:
            QMessageBox.information(self,'Pilih kandidat','Pilih baris kandidat terlebih dahulu.'); return
        self._active_table=self.auto_table
        self._activate_map_tool(fid, self.auto_table)

    def arm_selected_row(self):
        if not self.source or not hasattr(self, 'table') or self.table.currentRow() < 0: return
        fid=self.selected_fid(self.table)
        if fid is not None:
            self._active_table=self.table
            self._activate_map_tool(fid, self.table)

    def _flag_index(self):
        if not self.source:
            return -1
        return self._field_index([fld.name() for fld in self.source.fields()], ('flag',))

    def _flag_is_one(self, feature):
        idx = self._flag_index()
        if idx < 0 or feature is None:
            return False
        try:
            value = feature[idx]
            return str(value).strip() == '1' or (isinstance(value, (int, float)) and int(value) == 1)
        except Exception:
            return False

    def _zoom_to_flagged_feature(self, fid):
        """Zoom to the original point only when its source flag attribute equals 1."""
        if not self.source or self._flag_index() < 0:
            return False
        feature = self.source.getFeature(int(fid))
        if feature is None or not self._flag_is_one(feature) or not feature.hasGeometry() or feature.geometry().isEmpty():
            return False
        try:
            geom = feature.geometry()
            if QgsWkbTypes.isMultiType(geom.wkbType()):
                pts = geom.asMultiPoint()
                if not pts:
                    return False
                point = QgsPointXY(pts[0])
            else:
                point = QgsPointXY(geom.asPoint())
            canvas = self.iface.mapCanvas()
            dest_crs = canvas.mapSettings().destinationCrs()
            point = QgsCoordinateTransform(self.source.crs(), dest_crs, QgsProject.instance()).transform(point)
            current = canvas.extent()
            width = max(current.width() / 8.0, 0.0001)
            height = max(current.height() / 8.0, 0.0001)
            rect = QgsRectangle(point.x() - width/2.0, point.y() - height/2.0, point.x() + width/2.0, point.y() + height/2.0)
            canvas.setExtent(rect)
            canvas.refresh()
            self.status_label.setText(f'FLAG = 1 • Zoom otomatis ke titik FID {fid}. Setelah posisi dikoreksi, flag pada layer hasil menjadi 0.')
            return True
        except Exception as e:
            self.status_label.setText(f'Zoom flag gagal untuk FID {fid}: {e}')
            return False

    def _zoom_selected_auto_flag(self):
        if not self.source or not hasattr(self, 'auto_table'):
            return
        fid = self.selected_fid(self.auto_table)
        if fid is not None:
            self._zoom_to_flagged_feature(fid)

    def _activate_map_tool(self, fid, table):
        if not self.source: return
        self._active_fid=int(fid); self._active_table=table; self._pick_table=table
        self._zoom_to_flagged_feature(fid)
        if self._click_tool is None:
            self._old_tool=self.iface.mapCanvas().mapTool()
            self._click_tool=QgsMapToolEmitPoint(self.iface.mapCanvas())
            self._click_tool.canvasClicked.connect(self.on_canvas_click)
            self.iface.mapCanvas().setMapTool(self._click_tool)
        row=self._row_for_fid(table,fid); no=table.item(row,1).text() if row>=0 else '?'
        self.status_label.setText(f'KLIK PETA AKTIF • FID {fid} / no_bang {no}. Klik noktah hitam kecil. Setelah itu baris berikutnya otomatis aktif.')
        self._update_target_card(fid)
        self.iface.mapCanvas().setFocus()

    def on_canvas_click(self, point, button):
        if not self.source or self._active_fid is None: return
        try:
            tr=QgsCoordinateTransform(self.iface.mapCanvas().mapSettings().destinationCrs(),self.source.crs(),QgsProject.instance())
            p=tr.transform(point); fid=int(self._active_fid); table=self._pick_table or self.table
            self.mapping[fid]=QgsPointXY(p)
            active_fids=self._active_group_fids() or set(self._feature_records.keys())
            mapped_group=sum(1 for mapped_fid in self.mapping if mapped_fid in active_fids)
            pct=int(mapped_group*100/max(1,len(active_fids))); self.process_progress.setValue(pct); self.circular_progress.setValue(pct)
            self.process_progress.setFormat(f'PEMETAAN • {mapped_group}/{len(active_fids)}')
            for tbl in (self.table, self.auto_table):
                row = self._row_for_fid(tbl, fid)
                if row >= 0:
                    if tbl is self.table:
                        # Main table columns: ... flag, Status, Koordinat target
                        tbl.item(row, 5).setText('Target manual')
                        tbl.item(row, 6).setText(f'{p.x():.8f}, {p.y():.8f}')
                    else:
                        # Auto-review table columns: FID, no_bang, Flag hasil, Status, Koordinat target
                        tbl.item(row, 3).setText('Target manual (koreksi)')
                        tbl.item(row, 4).setText(f'{p.x():.8f}, {p.y():.8f}')
            # Keep map tool active and automatically arm next unmapped feature.
            self._active_fid=None
            if table is self.table:
                row=self._row_for_fid(self.table,fid); nextrow=-1
                for nxt in range(max(0,row+1),self.table.rowCount()):
                    nf=int(self.table.item(nxt,0).text())
                    if nf not in self.mapping: nextrow=nxt; break
                if nextrow < 0:
                    for nxt in range(self.table.rowCount()):
                        nf=int(self.table.item(nxt,0).text())
                        if nf not in self.mapping: nextrow=nxt; break
                if nextrow >= 0:
                    self.table.selectRow(nextrow); self.table.scrollToItem(self.table.item(nextrow,0))
                    self._active_fid=int(self.table.item(nextrow,0).text()); self._pick_table=self.table
                    nr=self.table.item(nextrow,1).text()
                    self.status_label.setText(f'Target {fid} tersimpan. Nomor berikutnya {nr} sudah aktif; klik noktah berikutnya atau pilih nomor lain.')
                else:
                    self.status_label.setText('Semua fitur sudah dipetakan. Buka Review untuk memeriksa perubahan di peta.')
                    self._active_fid=None
            else:
                self.status_label.setText(f'Target kandidat FID {fid} dikoreksi manual. Pilih kandidat lain atau buka Review.')
            self._update_target_card(self._active_fid)
            self.iface.mapCanvas().refresh()
        except Exception as e:
            QMessageBox.critical(self,'Gagal membaca koordinat',str(e))

    def stop_pick(self, restore=True):
        if self._click_tool:
            try: self._click_tool.canvasClicked.disconnect(self.on_canvas_click)
            except Exception: pass
            self._click_tool=None
        self._active_fid=None
        if restore and self._old_tool: self.iface.mapCanvas().setMapTool(self._old_tool)
        self._old_tool=None

    def select_next_unmapped(self):
        for row in range(self.table.rowCount()):
            fid=int(self.table.item(row,0).text())
            if fid not in self.mapping:
                self.table.selectRow(row); self.table.scrollToItem(self.table.item(row,0)); return
        QMessageBox.information(self,'Selesai','Semua fitur sudah dipetakan. Klik tombol Selesai & buat layer hasil.')

    def clear_selected(self):
        fid=self.selected_fid(self.table)
        if fid is None:return
        self.mapping.pop(fid,None); row=self._row_for_fid(self.table,fid)
        self.table.item(row,2).setText('Belum dipetakan'); self.table.item(row,3).setText('—')

    def _find_raster(self):
        # Use the explicitly selected raster; silently choosing the wrong raster was a common source of no candidates.
        if getattr(self, 'raster_combo', None) is not None:
            rid=self.raster_combo.currentData(); lyr=QgsProject.instance().mapLayer(rid) if rid else None
            if isinstance(lyr,QgsRasterLayer) and lyr.isValid() and not self._is_web_raster(lyr): return lyr
            # Web basemaps frequently return transparent/no-data blocks; prefer georeferenced local raster.
            local=[r for r in QgsProject.instance().mapLayers().values() if isinstance(r,QgsRasterLayer) and r.isValid() and not self._is_web_raster(r)]
            if local:
                local.sort(key=lambda r: ('wss' not in r.name().lower(), not self._layer_is_visible(r), r.name().lower()))
                return local[0]
            if isinstance(lyr,QgsRasterLayer) and lyr.isValid(): return lyr
        rasters=[l for l in QgsProject.instance().mapLayers().values() if isinstance(l,QgsRasterLayer) and l.isValid()]
        if not rasters:return None
        rasters.sort(key=lambda r:(self._is_web_raster(r),'wss' not in r.name().lower(),not self._layer_is_visible(r)))
        return rasters[0]

    def _active_group_fids(self):
        if not self._feature_group_values:
            return None
        village = self.group_primary_combo.currentData() if hasattr(self, 'group_primary_combo') else None
        sls = self.group_sls_combo.currentData() if hasattr(self, 'group_sls_combo') else None
        if village is None or sls is None:
            return set()
        return {int(fid) for fid, values in self._feature_group_values.items()
                if values.get('desa_code') == village and values.get('sls', '') == sls}

    def run_auto_suggestions(self):
        """OCR-match raster labels to no_bang; never use dark pixels as the target rule."""
        if not self.source:
            QMessageBox.warning(self, 'Layer belum dipilih', 'Pilih layer titik bangunan yang memiliki atribut no_bang terlebih dahulu.')
            return
        raster = self._find_raster()
        if raster is None:
            QMessageBox.warning(self, 'Raster tidak ditemukan', 'Muat raster georeferensi yang valid, lalu pilih pada daftar Raster referensi.')
            return
        if not raster.isValid() or not raster.crs().isValid():
            QMessageBox.warning(self, 'Raster tidak valid', 'Raster harus valid dan memiliki CRS yang benar.')
            return
        try:
            import pytesseract
        except Exception as import_error:
            QMessageBox.warning(
                self, 'Modul pytesseract belum tersedia',
                'Fitur OCR menggunakan Tesseract OCR dan modul Python pytesseract.\n\n'
                'Pasang modul pytesseract pada Python yang dipakai QGIS, lalu coba kembali.\n'
                f'Detail: {import_error}'
            )
            return
        try:
            tesseract_path, tesseract_version = _configure_tesseract_ocr(pytesseract)
        except Exception as ocr_error:
            QMessageBox.warning(
                self, 'Tesseract OCR belum siap',
                'Modul pytesseract sudah dimuat, tetapi aplikasi Tesseract OCR belum bisa dijalankan.\n\n'
                'Instal Tesseract OCR untuk Windows. Plugin mencari lokasi standar secara otomatis. '
                'Jika dipasang di lokasi khusus, atur variabel lingkungan TESSERACT_CMD ke path tesseract.exe.\n\n'
                f'Detail: {ocr_error}'
            )
            return
        try:
            from qgis.PyQt.QtWidgets import QApplication
            from qgis.PyQt.QtGui import QImage
            import statistics
            self._auto_candidates = {}
            active_fids = self._active_group_fids()
            feats = [f for f in self.source.getFeatures()
                     if active_fids is None or int(f.id()) in active_fids]
            if not feats:
                QMessageBox.information(self, 'Kelompok kosong', 'Kelompok desa/SLS yang dipilih tidak memiliki fitur untuk dianalisis.')
                return
            fields = [field.name().casefold() for field in self.source.fields()]
            no_idx = next((i for i, name in enumerate(fields)
                           if name in ('no_bang', 'nobang', 'no bang', 'nomor_bangunan')), -1)
            if no_idx < 0:
                QMessageBox.warning(self, 'Atribut no_bang tidak ditemukan',
                                    'Layer titik bangunan harus memiliki kolom no_bang.')
                return
            provider = raster.dataProvider()
            if provider.bandCount() < 1:
                QMessageBox.warning(self, 'Raster tidak terbaca', 'Raster tidak memiliki band yang dapat dibaca.')
                return
            radius_m = float(self.radius.value())
            side = max(180, min(640, int(self.samples.value()) * 16))
            scale = 3
            project = QgsProject.instance()
            to_raster = QgsCoordinateTransform(self.source.crs(), raster.crs(), project)
            to_source = QgsCoordinateTransform(raster.crs(), self.source.crs(), project)
            extent = raster.extent()
            width_total = max(1, provider.xSize())
            height_total = max(1, provider.ySize())
            gt = None
            try:
                from osgeo import gdal
                ds = gdal.Open(raster.source().split('|')[0])
                if ds:
                    gt = ds.GetGeoTransform()
                    ds = None
            except Exception:
                gt = None

            def normalized_number(value):
                digits = re.sub(r'\D', '', str(value or ''))
                if not digits:
                    return None
                try:
                    return str(int(digits))
                except Exception:
                    return digits.lstrip('0') or '0'

            def distance_m(p1, p2):
                # Distances are always reported and compared in metres, including EPSG:4326.
                dx = p1.x() - p2.x()
                dy = p1.y() - p2.y()
                if raster.crs().mapUnits() == QgsUnitTypes.DistanceDegrees:
                    lat = (p1.y() + p2.y()) / 2.0
                    east = dx * 111320.0 * max(0.01, math.cos(math.radians(lat)))
                    north = dy * 111320.0
                    return math.hypot(east, north)
                try:
                    factor = QgsUnitTypes.fromUnitToUnitFactor(raster.crs().mapUnits(), QgsUnitTypes.DistanceMeters)
                    factor = factor if factor and factor > 0 else 1.0
                except Exception:
                    factor = 1.0
                return math.hypot(dx, dy) * factor

            def search_rect(point):
                # Radius in metres; convert to map units for geographic or projected raster CRS.
                if raster.crs().mapUnits() == QgsUnitTypes.DistanceDegrees:
                    lat = max(-89.0, min(89.0, point.y()))
                    dy = radius_m / 111320.0
                    dx = radius_m / max(1000.0, 111320.0 * math.cos(math.radians(lat)))
                else:
                    try:
                        factor = QgsUnitTypes.fromUnitToUnitFactor(raster.crs().mapUnits(), QgsUnitTypes.DistanceMeters)
                        factor = factor if factor and factor > 0 else 1.0
                    except Exception:
                        factor = 1.0
                    dx = dy = radius_m / factor
                return QgsRectangle(point.x()-dx, point.y()-dy, point.x()+dx, point.y()+dy)

            total = max(1, len(feats))
            self._auto_candidate_meta = {}  # fid -> (normalized label, OCR map point, distance metres, confidence)
            self.auto_progress.setValue(0); self.circular_progress.setValue(0)
            self.auto_progress.setFormat('Mencocokkan nomor • %p%')
            self.auto_btn.setEnabled(False); self.auto_btn.setText('Membaca nomor…')
            QApplication.setOverrideCursor(Qt.WaitCursor)
            temp_png = os.path.join(tempfile.gettempdir(), 'slsmagic_ocr_crop.png')
            for i, feat in enumerate(feats):
                fid = int(feat.id())
                row = self._row_for_fid(self.auto_table, fid)
                try:
                    raw_no = feat.attribute(no_idx)
                    target_no = normalized_number(raw_no)
                    if target_no is None:
                        if row >= 0:
                            self.auto_table.item(row, 5).setText('no_bang kosong')
                            self.auto_table.item(row, 6).setText('—')
                        continue
                    geom = feat.geometry()
                    if geom is None or geom.isEmpty():
                        if row >= 0:
                            self.auto_table.item(row, 5).setText('Geometri kosong')
                        continue
                    src_pt = geom.asPoint() if QgsWkbTypes.geometryType(geom.wkbType()) == QgsWkbTypes.PointGeometry else geom.pointOnSurface().asPoint()
                    raster_pt = to_raster.transform(QgsPointXY(src_pt))
                    rect = search_rect(raster_pt)
                    # Restrict the search to raster coverage.
                    rect = rect.intersect(extent)
                    if rect.isEmpty():
                        if row >= 0:
                            self.auto_table.item(row, 5).setText('Di luar raster')
                        continue
                    w = h = side
                    # Use RGB luminance for multi-band imagery; band 1 alone may be a color channel.
                    band_count = raster.bandCount()
                    band_ids = [1, 2, 3] if band_count >= 3 else [1]
                    blocks = [provider.block(band, rect, w, h) for band in band_ids]
                    if not blocks or any(b is None or b.isEmpty() for b in blocks):
                        if row >= 0:
                            self.auto_table.item(row, 5).setText('Raster tidak terbaca')
                        continue
                    def _gray_at(yy, xx):
                        channels = [float(b.value(yy, xx)) for b in blocks]
                        if not all(math.isfinite(v) for v in channels):
                            return None
                        if len(channels) >= 3:
                            return 0.299 * channels[0] + 0.587 * channels[1] + 0.114 * channels[2]
                        return channels[0]
                    values = []
                    for yy in range(h):
                        for xx in range(w):
                            try:
                                val = _gray_at(yy, xx)
                                if val is not None:
                                    values.append(val)
                            except Exception:
                                pass
                    if not values:
                        if row >= 0:
                            self.auto_table.item(row, 5).setText('Tidak ada piksel valid')
                        continue
                    ordered = sorted(values)
                    lo = ordered[int((len(ordered)-1)*0.02)]
                    hi = ordered[int((len(ordered)-1)*0.98)]
                    if hi <= lo:
                        hi = lo + 1.0
                    buf = bytearray(w*h)
                    for yy in range(h):
                        for xx in range(w):
                            try:
                                val = _gray_at(yy, xx)
                                if val is None:
                                    val = lo
                                pix = int(max(0, min(255, (val-lo)*255.0/(hi-lo))))
                                buf[yy*w+xx] = pix
                            except Exception:
                                buf[yy*w+xx] = 255
                    image = QImage(bytes(buf), w, h, w, QImage.Format_Grayscale8).copy()
                    image = image.scaled(w*scale, h*scale)
                    if not image.save(temp_png, 'PNG'):
                        raise RuntimeError('Gagal membuat gambar sementara untuk OCR.')
                    data = pytesseract.image_to_data(
                        temp_png, config='--psm 11 -c tessedit_char_whitelist=0123456789',
                        output_type=pytesseract.Output.DICT)
                    matches = []
                    for j, txt in enumerate(data.get('text', [])):
                        got = normalized_number(txt)
                        if got != target_no:
                            continue
                        try:
                            conf = float(data.get('conf', ['-1'])[j])
                        except Exception:
                            conf = -1
                        if conf < 0:
                            continue
                        left = float(data['left'][j]) / scale
                        top = float(data['top'][j]) / scale
                        bw = float(data['width'][j]) / scale
                        bh = float(data['height'][j]) / scale
                        # Convert OCR label centre from crop pixels back to map coordinates.
                        mx = rect.xMinimum() + ((left + bw/2.0)/w) * rect.width()
                        my = rect.yMaximum() - ((top + bh/2.0)/h) * rect.height()
                        ocr_pt = QgsPointXY(mx, my)
                        dist = distance_m(ocr_pt, raster_pt)
                        matches.append((dist, -conf, ocr_pt, txt, conf))
                    if matches:
                        matches.sort(key=lambda item: (item[0], item[1]))
                        best = matches[0]
                        self._auto_candidates[fid] = to_source.transform(best[2])
                        self._auto_candidate_meta[fid] = (target_no, QgsPointXY(best[2]), float(best[0]), float(best[4]))
                        if row >= 0:
                            self.auto_table.item(row, 5).setText(f'Nomor cocok: {best[3]} • review')
                            self.auto_table.item(row, 6).setText(f'OCR {best[4]:.0f}% • {best[0]:.2f} m')
                    else:
                        if row >= 0:
                            self.auto_table.item(row, 5).setText('Nomor tidak ditemukan')
                            self.auto_table.item(row, 6).setText('—')
                except Exception as feature_error:
                    if row >= 0:
                        self.auto_table.item(row, 5).setText('Gagal analisis')
                        self.auto_table.item(row, 6).setText(f'{type(feature_error).__name__}: {feature_error}'[:100])
                finally:
                    pct = int((i+1)*100/total)
                    self.auto_progress.setValue(pct); self.circular_progress.setValue(pct)
            # One-to-one guard: if multiple point features selected the same OCR label,
            # Duplicate point features may legitimately share one raster label (e.g. 01/001/0001).
            # Keep every match, but fan matching points out by a small metric offset so they do not overlap.
            candidate_ids = list(self._auto_candidate_meta.keys())
            candidate_ids.sort(key=lambda fid: (self._auto_candidate_meta[fid][2], -self._auto_candidate_meta[fid][3]))
            clusters = []  # {number, raster label point, feature ids}
            for fid in candidate_ids:
                label_no, label_pt, dist_m, conf = self._auto_candidate_meta[fid]
                found = None
                for cluster in clusters:
                    if cluster['number'] == label_no and distance_m(label_pt, cluster['label_pt']) <= 1.5:
                        found = cluster
                        break
                if found is None:
                    found = {'number': label_no, 'label_pt': label_pt, 'fids': []}
                    clusters.append(found)
                found['fids'].append(fid)
            try:
                project = QgsProject.instance()
                metric_crs = QgsCoordinateReferenceSystem('EPSG:3857')
                to_metric = QgsCoordinateTransform(self.source.crs(), metric_crs, project)
                from_metric = QgsCoordinateTransform(metric_crs, self.source.crs(), project)
                for cluster in clusters:
                    fids = cluster['fids']
                    if len(fids) < 2:
                        continue
                    # Small ring offsets: first point is also offset so all duplicates remain distinguishable.
                    radius_offset_m = 0.30
                    for idx, fid in enumerate(fids):
                        point = QgsPointXY(self._auto_candidates[fid])
                        metric_point = to_metric.transform(point)
                        angle = (2.0 * math.pi * idx) / len(fids)
                        shifted = QgsPointXY(metric_point.x() + radius_offset_m * math.cos(angle),
                                             metric_point.y() + radius_offset_m * math.sin(angle))
                        self._auto_candidates[fid] = from_metric.transform(shifted)
                        row = self._row_for_fid(self.auto_table, fid)
                        if row >= 0 and len(fids) > 1:
                            self.auto_table.item(row, 5).setText('Nomor cocok • posisi digeser agar tidak tumpang tindih')
                            self.auto_table.item(row, 6).setText(f'Duplikat label {cluster['number']} • offset {radius_offset_m:.2f} m')
            except Exception as offset_error:
                # Keep candidates rather than dropping valid matches if coordinate transformation fails.
                self.status_label.setText(f'Peringatan offset duplikat: {offset_error}')

            self.auto_progress.setFormat('Pencocokan nomor selesai • %p%')
            self.status_label.setText(
                f'OCR Tesseract {tesseract_version} ({tesseract_path}) • '
                f'Nomor cocok pada {len(self._auto_candidates)} dari {len(feats)} fitur. '
                'Label raster satu digit dicocokkan secara numerik: no_bang 001/0001 sama dengan label 1. '
                'Nomor duplikat tetap dicocokkan; titik duplikat diberi offset kecil agar tidak bertumpuk. Tinjau posisi sebelum diterapkan.'
            )
            if not self._auto_candidates:
                QMessageBox.information(self, 'Tidak ada nomor cocok',
                    'Tidak ditemukan nomor yang cocok. Pastikan Tesseract OCR terpasang, radius pencarian cukup, dan resolusi raster memadai.')
        except Exception as e:
            import traceback
            details = traceback.format_exc()
            self.status_label.setText(f'Pencocokan nomor gagal: {type(e).__name__}: {e}')
            QMessageBox.critical(self, 'Pencocokan nomor gagal', f'{type(e).__name__}: {e}\n\nDetail teknis:\n{details[-1800:]}')
        finally:
            try:
                QApplication.restoreOverrideCursor()
            except Exception:
                pass
            self.auto_btn.setEnabled(True); self.auto_btn.setText('🔎 Cari saran otomatis')

    def apply_selected_suggestions(self):
        fid = self.selected_fid(self.auto_table)
        if fid is None:
            QMessageBox.information(self, 'Pilih kandidat', 'Pilih satu baris kandidat terlebih dahulu.')
            return
        if fid not in self._auto_candidates:
            QMessageBox.information(self, 'Tidak ada saran', 'Nomor yang cocok belum ditemukan untuk fitur ini.')
            return
        self.mapping[fid] = QgsPointXY(self._auto_candidates[fid])
        p = self.mapping[fid]
        row = self._row_for_fid(self.auto_table, fid)
        if row >= 0:
            self.auto_table.item(row, 5).setText('Saran diterapkan • periksa')
            self.auto_table.item(row, 6).setText(f'{p.x():.8f}, {p.y():.8f}')
        mr = self._row_for_fid(self.table, fid)
        if mr >= 0:
            self.table.item(mr, 5).setText('Target otomatis (periksa)')
            self.table.item(mr, 6).setText(f'{p.x():.8f}, {p.y():.8f}')
        self.status_label.setText(
            f'Saran nomor diterapkan untuk FID {fid}. Periksa posisi titik pada peta sebelum membuat layer hasil.')

    def open_review(self):
        if not self.source: return
        self.review_dialog=ReviewDialog(self)
        self.review_dialog.setWindowModality(Qt.NonModal)
        self.review_dialog.show(); self.review_dialog.raise_(); self.review_dialog.activateWindow()

    def finish_to_layer(self):
        """Write a durable GeoPackage output; never add a temporary memory result."""
        if not self.source:
            return
        active_fids = self._active_group_fids()
        if not active_fids:
            QMessageBox.information(self, 'Pilih SLS', 'Pilih kode desa di sebelah kiri dan nama SLS di sebelah kanan sebelum membuat layer hasil.')
            return
        group_mapping = {fid: point for fid, point in self.mapping.items() if fid in active_fids}
        path, _ = QFileDialog.getSaveFileName(self, 'Simpan hasil Auto Snap sebagai GeoPackage',
            os.path.join(os.path.expanduser('~'), f'AutoSnap_{re.sub(r"[^A-Za-z0-9_-]+", "_", str(self._selected_sls_name or "SLS"))}.gpkg'),
            'GeoPackage (*.gpkg)')
        if not path:
            return
        if not path.lower().endswith('.gpkg'):
            path += '.gpkg'
        answer = QMessageBox.question(self, 'Konfirmasi output GPKG',
            f'Simpan hasil permanen ke:\n{path}\n\nSLS: {self._selected_sls_name}\nFitur sumber: {self.source.featureCount()}\nTitik yang dipindahkan: {len(group_mapping)}\n\nLayer sumber tidak diubah.',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if answer != QMessageBox.Yes:
            return
        try:
            safe_sls = re.sub(r'[^A-Za-z0-9_-]+', '_', str(self._selected_sls_name or 'SLS')).strip('_')[:48]
            out_name = f'AutoSnap_{self._selected_village_code}_{safe_sls}'[:60]
            out = QgsVectorLayer(f'{QgsWkbTypes.displayString(self.source.wkbType())}?crs={self.source.crs().authid()}', out_name, 'memory')
            fields = list(self.source.fields()); existing = [f.name().lower() for f in fields]
            add_desa = 'nama_desa' not in existing; add_sls = 'nama_sls' not in existing; add_flag = 'flag' not in existing
            if add_desa: fields.append(QgsField('nama_desa', QVariant.String, len=120))
            if add_sls: fields.append(QgsField('nama_sls', QVariant.String, len=120))
            if add_flag: fields.append(QgsField('flag', QVariant.Int))
            dp = out.dataProvider(); dp.addAttributes(fields); out.updateFields(); new = []
            names = [fld.name() for fld in self.source.fields()]
            flag_idx = self._field_index(names, ('flag',))
            di = self._field_index(names, ('nama_desa','desa','desa_name','village','nama_kelurahan','kelurahan'))
            si = self._field_index(names, ('nama_sls','sls','level_6_name','sls_name','nama_satuan_lingkungan_setempat'))
            desa_ui = self.desa_edit.text().strip() if hasattr(self, 'desa_edit') else ''
            sls_ui = self.sls_edit.text().strip() if hasattr(self, 'sls_edit') else ''
            for f in self.source.getFeatures():
                nf = QgsFeature(out.fields()); attrs = list(f.attributes())
                desa_val = str(f[di]).strip() if di >= 0 and f[di] not in (None, '') else desa_ui
                sls_val = str(f[si]).strip() if si >= 0 and f[si] not in (None, '') else sls_ui
                if add_desa: attrs.append(desa_val)
                elif di >= 0 and (f[di] in (None, '') or not str(f[di]).strip()): attrs[di] = desa_ui
                if add_sls: attrs.append(sls_val)
                elif si >= 0 and (f[si] in (None, '') or not str(f[si]).strip()): attrs[si] = sls_ui
                if flag_idx >= 0 and int(f.id()) in group_mapping: attrs[flag_idx] = 0
                if add_flag: attrs.append(0 if int(f.id()) in group_mapping else 1)
                nf.setAttributes(attrs)
                geom = QgsGeometry(f.geometry())
                if int(f.id()) in group_mapping: geom = QgsGeometry.fromPointXY(group_mapping[int(f.id())])
                nf.setGeometry(geom); new.append(nf)
            dp.addFeatures(new); out.updateExtents()
            opts = QgsVectorFileWriter.SaveVectorOptions(); opts.driverName = 'GPKG'; opts.layerName = out_name
            if os.path.exists(path):
                # Avoid overwriting unrelated datasets; user explicitly selected the target path.
                opts.actionOnExistingFile = QgsVectorFileWriter.CreateOrOverwriteLayer
            result = QgsVectorFileWriter.writeAsVectorFormatV3(out, path, QgsProject.instance().transformContext(), opts)
            if result[0] != QgsVectorFileWriter.NoError:
                raise RuntimeError(str(result))
            saved = QgsVectorLayer(path + '|layername=' + out_name, out_name, 'ogr')
            if not saved.isValid():
                raise RuntimeError('GPKG berhasil ditulis tetapi layer hasil tidak dapat dibuka kembali.')
            QgsProject.instance().addMapLayer(saved)
            self.apply_style(saved, self._distinct_output_color())
            self.status_label.setText(f'Output permanen tersimpan: {path}. Titik dipindahkan: {len(group_mapping)}. Layer sumber tetap tidak berubah.')
            QMessageBox.information(self, 'Output GPKG berhasil', f'GeoPackage tersimpan dan dimuat ke proyek.\n\nFile: {path}\nLayer: {out_name}\nTotal fitur: {saved.featureCount()}\nTitik yang dipindahkan: {len(group_mapping)}\n\nLayer output bukan temporary layer.')
        except Exception as e:
            QMessageBox.critical(self, 'Gagal menyimpan GPKG', f'{type(e).__name__}: {e}')

    def _distinct_output_color(self):
        """Pick one coherent color per output layer, distinct from the default green style."""
        palette=['235,95,55,255','45,112,220,255','153,91,204,255','224,151,35,255','20,154,164,255','210,65,125,255','92,104,220,255']
        used=[]
        for lyr in QgsProject.instance().mapLayers().values():
            if lyr.name().startswith('Titik Bangunan • AutoSnap'):
                try:
                    renderer=lyr.renderer(); sym=renderer.symbol() if renderer else None
                    if sym: used.append(sym.color().name())
                except Exception: pass
        available=[QColor(int(c.split(',')[0]),int(c.split(',')[1]),int(c.split(',')[2])).name() for c in palette]
        unused=[i for i,c in enumerate(available) if c not in used]
        idx=random.choice(unused or list(range(len(palette))))
        return palette[idx]

    def apply_style(self, layer=None, color=None):
        layer=layer or self.source
        if not layer: return
        # Exactly one symbol color for every feature in this layer; never per-FID categories.
        color = color or QColor(70,145,105)
        if isinstance(color, QColor): color_string=f'{color.red()},{color.green()},{color.blue()},255'
        else: color_string=str(color)
        symbol=QgsMarkerSymbol.createSimple({'name':'circle','color':color_string,'outline_color':'28,55,42,255','outline_width':'0.25','size':'4'})
        layer.setRenderer(QgsSingleSymbolRenderer(symbol))
        pal=QgsPalLayerSettings(); pal.fieldName='no_bang'; pal.enabled=True
        fmt=QgsTextFormat(); font=QFont('Arial',10); font.setBold(True); fmt.setFont(font); fmt.setSize(10); fmt.setColor(QColor('white'))
        buf=QgsTextBufferSettings(); buf.setEnabled(True); buf.setSize(1.2); buf.setColor(QColor('black')); fmt.setBuffer(buf)
        shadow=QgsTextShadowSettings(); shadow.setEnabled(True); shadow.setColor(QColor('black')); shadow.setBlurRadius(1.5); shadow.setOffsetDistance(1.0); fmt.setShadow(shadow)
        pal.setFormat(fmt); layer.setLabeling(QgsVectorLayerSimpleLabeling(pal)); layer.setLabelsEnabled(True); layer.triggerRepaint()


    def closeEvent(self, event):
        try:
            self.stop_pick(restore=True)
        except Exception:
            pass
        try:
            self.restore_source_subset()
        except Exception:
            pass
        try:
            self.restore_georeference_visibility()
        except Exception:
            pass
        super().closeEvent(event)

class ReviewDialog(QDialog):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner=owner
        self.setWindowTitle('SLS MAGIC • Review hasil AutoSnap')
        self.resize(820,500); self.setMinimumSize(600,360); self.setSizeGripEnabled(True)
        lay=QVBoxLayout(self); lay.setContentsMargins(12,12,12,12); lay.setSpacing(8)
        lay.addWidget(QLabel('<b>Review sebelum membuat layer baru</b> — periksa nomor, status, dan koordinat target. Belum ada data yang disimpan ke disk.'))
        self.table=QTableWidget(0,5); self.table.setHorizontalHeaderLabels(['FID','no_bang','Flag hasil','Status','Koordinat target'])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows); self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(0,QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1,QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2,QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3,QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4,QHeaderView.Stretch)
        lay.addWidget(self.table,1)
        bar=QHBoxLayout(); self.refresh_btn=QPushButton('Perbarui review'); self.refresh_btn.clicked.connect(self.refresh)
        self.finish_btn=QPushButton('SELESAI • Buat layer baru'); self.finish_btn.clicked.connect(self.finish)
        close=QPushButton('Kembali'); close.clicked.connect(self.close)
        bar.addWidget(self.refresh_btn); bar.addStretch(1); bar.addWidget(close); bar.addWidget(self.finish_btn); lay.addLayout(bar)
        self.refresh()
    def create_preview_layer(self):
        src=self.owner.source
        if not src: return
        try:
            for old in (self.owner._preview_layer, self.owner._changed_layer):
                if old and QgsProject.instance().mapLayer(old.id()): QgsProject.instance().removeMapLayer(old.id())
            self.owner._preview_layer=None; self.owner._changed_layer=None
            name='AutoSnap • Review Preview'
            for lyr in list(QgsProject.instance().mapLayers().values()):
                if lyr.name()==name: QgsProject.instance().removeMapLayer(lyr.id())
            mem=QgsVectorLayer(f'{QgsWkbTypes.displayString(src.wkbType())}?crs={src.crs().authid()}',name,'memory')
            fields=list(src.fields()); existing=[fld.name().lower() for fld in fields]
            add_desa='nama_desa' not in existing; add_sls='nama_sls' not in existing; add_flag='flag' not in existing
            if add_desa: fields.append(QgsField('nama_desa', QVariant.String, len=120))
            if add_sls: fields.append(QgsField('nama_sls', QVariant.String, len=120))
            if add_flag: fields.append(QgsField('flag', QVariant.Int))
            dp=mem.dataProvider(); dp.addAttributes(fields); mem.updateFields(); new=[]
            names=[fld.name() for fld in src.fields()]
            flag_idx=self.owner._field_index(names, ('flag',))
            di=self.owner._field_index(names, ('nama_desa','desa','desa_name','village','nama_kelurahan','kelurahan'))
            si=self.owner._field_index(names, ('nama_sls','sls','level_6_name','sls_name','nama_satuan_lingkungan_setempat'))
            for f in src.getFeatures():
                nf=QgsFeature(mem.fields()); attrs=list(f.attributes())
                desa_value=str(f[di]).strip() if di>=0 and f[di] not in (None,'') else self.owner.desa_edit.text().strip()
                sls_value=str(f[si]).strip() if si>=0 and f[si] not in (None,'') else self.owner.sls_edit.text().strip()
                if add_desa: attrs.append(desa_value)
                elif di>=0 and (f[di] in (None,'') or str(f[di]).strip()==''): attrs[di]=desa_value
                if add_sls: attrs.append(sls_value)
                elif si>=0 and (f[si] in (None,'') or str(f[si]).strip()==''): attrs[si]=sls_value
                if flag_idx >= 0 and int(f.id()) in self.owner.mapping: attrs[flag_idx] = 0
                if add_flag: attrs.append(0 if int(f.id()) in self.owner.mapping else 1)
                nf.setAttributes(attrs); geom=QgsGeometry(f.geometry())
                if f.id() in self.owner.mapping: geom=QgsGeometry.fromPointXY(self.owner.mapping[f.id()])
                nf.setGeometry(geom); new.append(nf)
            dp.addFeatures(new); mem.updateExtents(); QgsProject.instance().addMapLayer(mem); self.owner.apply_style(mem, self.owner._distinct_output_color()); self.owner._preview_layer=mem
            # Overlay only moved features with a larger contrasting symbol so review visibly shows edits.
            changed=QgsVectorLayer(f'{QgsWkbTypes.displayString(src.wkbType())}?crs={src.crs().authid()}','AutoSnap • Titik Berubah','memory')
            cdp=changed.dataProvider(); cdp.addAttributes(list(src.fields())); changed.updateFields(); moved=[]
            for f in src.getFeatures():
                if f.id() not in self.owner.mapping: continue
                cf=QgsFeature(changed.fields()); cf.setAttributes(f.attributes()); cf.setGeometry(QgsGeometry.fromPointXY(self.owner.mapping[f.id()])); moved.append(cf)
            cdp.addFeatures(moved); changed.updateExtents(); QgsProject.instance().addMapLayer(changed)
            highlight=QgsMarkerSymbol.createSimple({'name':'circle','color':'255,65,54,255','outline_color':'255,255,255,255','outline_width':'0.7','size':'5.5'})
            changed.setRenderer(QgsSingleSymbolRenderer(highlight)); changed.setLabelsEnabled(False); self.owner._changed_layer=changed
            self.owner.iface.mapCanvas().refresh()
        except Exception as e:
            QMessageBox.warning(self,'Preview peta tidak tersedia',str(e))
    def closeEvent(self,event):
        try:
            lyr=self.owner._preview_layer
            if lyr and QgsProject.instance().mapLayer(lyr.id()): QgsProject.instance().removeMapLayer(lyr.id())
            self.owner._preview_layer=None
            if self.owner._changed_layer and QgsProject.instance().mapLayer(self.owner._changed_layer.id()): QgsProject.instance().removeMapLayer(self.owner._changed_layer.id())
            self.owner._changed_layer=None; self.owner.iface.mapCanvas().refresh()
        except Exception: pass
        super().closeEvent(event)
    def refresh(self):
        src=self.owner.source
        if not src:return
        feats=list(src.getFeatures()); feats.sort(key=lambda f:(str(f['no_bang']).strip().zfill(8),f.id()))
        self.table.setRowCount(len(feats))
        for r,f in enumerate(feats):
            p=self.owner.mapping.get(f.id())
            status='Target tersimpan' if p else 'Posisi asli (tidak diubah)'
            coord=f'{p.x():.8f}, {p.y():.8f}' if p else '—'
            names=[fld.name() for fld in src.fields()]; flag_idx=self.owner._field_index(names, ('flag',))
            if p: flag_value='0'
            elif flag_idx >= 0 and f[flag_idx] not in (None, ''): flag_value=str(f[flag_idx])
            else: flag_value='1'
            for c,v in enumerate([str(f.id()),str(f['no_bang']).strip(),flag_value,status,coord]): self.table.setItem(r,c,QTableWidgetItem(v))
        self.create_preview_layer()
    def finish(self):
        self.owner.finish_to_layer()
        if self.owner._preview_layer:
            try: QgsProject.instance().removeMapLayer(self.owner._preview_layer.id())
            except Exception: pass
            self.owner._preview_layer=None
        if self.owner._changed_layer:
            try: QgsProject.instance().removeMapLayer(self.owner._changed_layer.id())
            except Exception: pass
            self.owner._changed_layer=None
        self.close()

class PointAuditDialog(QDialog):
    'Modern audit screen: choose a reference point layer and a project group/layer.'
    def __init__(self, iface, parent=None):
        super().__init__(parent); self.iface = iface
        self.setWindowTitle('SLS MAGIC  |  Audit Laporan Titik')
        self.resize(900, 620); self.setMinimumSize(700, 500)
        lay = QVBoxLayout(self); lay.setContentsMargins(18, 16, 18, 16); lay.setSpacing(12)
        title = QLabel('AUDIT LAPORAN TITIK'); title.setStyleSheet('font-size:18pt;font-weight:800;color:#12304a;')
        subtitle = QLabel('Pilih layer acuan kecamatan dan folder pekerjaan untuk memantau progres titik bangunan.')
        subtitle.setWordWrap(True); subtitle.setStyleSheet('color:#607086;font-size:10pt;')
        lay.addWidget(title); lay.addWidget(subtitle)
        selector = QFrame(); selector.setObjectName('metaPanel'); sr = QVBoxLayout(selector); sr.setContentsMargins(12, 10, 12, 10); sr.setSpacing(6)
        sr.addWidget(QLabel('<b>1. Layer Acuan Kecamatan dengan Geotaging Kecamatan</b>'))
        row = QHBoxLayout(); self.reference = QComboBox()
        # Sumber acuan tetap disediakan bersama plugin agar titik acuan BPS Pusat
        # konsisten dan tidak bergantung pada layer aktif atau nama layer proyek.
        project = QgsProject.instance()
        fixed_name = 'KdKec_3210200 (Acuan Titik Pusat BPS)'
        fixed_layer = next((l for l in project.mapLayers().values()
            if isinstance(l, QgsVectorLayer) and l.isValid()
            and l.name() == fixed_name), None)
        if fixed_layer is None:
            fixed_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'KdKec_3210200.gpkg')
            if os.path.exists(fixed_path):
                fixed_layer = QgsVectorLayer(
                    fixed_path + '|layername=KdKec_3210200',
                    fixed_name, 'ogr')
                if fixed_layer.isValid() and fixed_layer.geometryType() == QgsWkbTypes.PointGeometry:
                    project.addMapLayer(fixed_layer, False)
                else:
                    fixed_layer = None

        all_point_layers = [l for l in project.mapLayers().values()
            if isinstance(l, QgsVectorLayer) and l.isValid()
            and l.geometryType() == QgsWkbTypes.PointGeometry]
        # Acuan bawaan diprioritaskan, tetapi dropdown tetap aktif agar pengguna
        # masih bisa mengganti pilihan secara manual bila memang diperlukan.
        def reference_priority(layer):
            name = layer.name().casefold()
            if name.startswith('kdkec') or 'kdkec_' in name:
                return 0
            if 'kecamatan' in name:
                return 1
            if 'geotag' in name or 'geotaging' in name:
                return 2
            return 3
        self.layers = sorted(all_point_layers, key=lambda layer: (reference_priority(layer), layer.name().casefold()))
        active = self.iface.activeLayer() if hasattr(self.iface, 'activeLayer') else None
        for l in self.layers:
            self.reference.addItem(l.name(), l.id())
        preferred = None
        if fixed_layer and any(l.id() == fixed_layer.id() for l in self.layers):
            preferred = fixed_layer
        elif active and any(l.id() == active.id() for l in self.layers):
            preferred = active
        elif self.layers:
            preferred = self.layers[0]
        if preferred:
            self.reference.setCurrentIndex(self.reference.findData(preferred.id()))
        self.reference.setMinimumHeight(36)
        self.reference.setEnabled(True)
        self.reference.setToolTip('Pilih layer titik acuan BPS Pusat. Disarankan menggunakan layer Kecamatan, misalnya KdKec_...')
        row.addWidget(self.reference, 1)
        if not self.layers:
            self.reference.addItem('Tidak ada layer titik yang valid di project', None)
            self.reference.setEnabled(False)
        refresh = QPushButton('↻  Hitung ulang')
        refresh.setObjectName('auditRefreshButton')
        refresh.setMinimumHeight(38)
        refresh.setToolTip('Hitung ulang statistik berdasarkan layer acuan dan grup pekerjaan yang dipilih.')
        refresh.clicked.connect(self.recount)
        row.addWidget(refresh)

        sr.addLayout(row); lay.addWidget(selector)
        cards = QHBoxLayout(); cards.setSpacing(10)
        self.total = QLabel('—'); self.result = QLabel('—'); self.layers_count = QLabel('—')
        def metric(caption, value, accent):
            card = QFrame(); card.setStyleSheet('QFrame{background:#ffffff;border:1px solid #dce5ef;border-radius:10px;}')
            v = QVBoxLayout(card); v.setContentsMargins(13,10,13,10); v.setSpacing(3)
            cap = QLabel(caption); cap.setStyleSheet('color:#64748b;font-size:9pt;font-weight:600;border:none;')
            value.setStyleSheet(f'color:{accent};font-size:22pt;font-weight:800;border:none;')
            v.addWidget(cap); v.addWidget(value); cards.addWidget(card, 1)
        metric('TOTAL TITIK BPS PUSAT', self.total, '#1769aa')
        metric('PROGRES TITIK BANGUNAN', self.result, '#16804a')
        metric('TOTAL LAYER', self.layers_count, '#7755b8')
        lay.addLayout(cards)
        progress_panel = QFrame(); progress_panel.setObjectName('coveragePanel')
        pp = QVBoxLayout(progress_panel); pp.setContentsMargins(14, 11, 14, 12); pp.setSpacing(7)
        progress_head = QHBoxLayout()
        self.coverage_title = QLabel('PROGRES MONITORING TITIK BANGUNAN FINAL')
        self.coverage_title.setStyleSheet('font-size:9pt;font-weight:800;color:#334155;border:none;')
        self.coverage_percent = QLabel('0%')
        self.coverage_percent.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.coverage_percent.setStyleSheet('font-size:20pt;font-weight:900;color:#16804a;border:none;')
        progress_head.addWidget(self.coverage_title, 1); progress_head.addWidget(self.coverage_percent)
        pp.addLayout(progress_head)
        self.coverage_bar = QProgressBar(); self.coverage_bar.setRange(0, 100); self.coverage_bar.setValue(0)
        self.coverage_bar.setFormat('%p%  •  progres terhadap total titik BPS Pusat')
        self.coverage_bar.setMinimumHeight(24); pp.addWidget(self.coverage_bar)
        self.coverage_detail = QLabel('Pilih folder/grup pekerjaan untuk menghitung progres.')
        self.coverage_detail.setWordWrap(True); self.coverage_detail.setStyleSheet('color:#64748b;font-size:9pt;border:none;')
        pp.addWidget(self.coverage_detail)
        lay.addWidget(progress_panel)
        lay.addWidget(QLabel('<b>2. Pilih grup/layer pekerjaan</b>'))
        hint = QLabel('Pilih folder pekerjaan untuk menghitung seluruh layer titik di dalam folder tersebut.')
        hint.setWordWrap(True); hint.setStyleSheet('color:#64748b;'); lay.addWidget(hint)
        self.groups = QTreeWidget(); self.groups.setHeaderLabels(['Grup / layer pekerjaan', 'Jumlah titik']); self.groups.setAlternatingRowColors(True)
        self.groups.setSelectionMode(QAbstractItemView.SingleSelection); self.groups.header().setSectionResizeMode(0, QHeaderView.Stretch); self.groups.header().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        lay.addWidget(self.groups, 1)
        root = QgsProject.instance().layerTreeRoot()
        def walk(node, parent_item=None):
            for child in node.children():
                if child.nodeType() == 0:
                    item = QTreeWidgetItem([child.name(), 'Group']); item.setData(0, Qt.UserRole, ('group', child.name()))
                    (parent_item.addChild(item) if parent_item else self.groups.addTopLevelItem(item))
                    walk(child, item)
                elif child.nodeType() == 1:
                    lyr = child.layer()
                    if lyr and isinstance(lyr, QgsVectorLayer) and lyr.isValid() and lyr.geometryType() == QgsWkbTypes.PointGeometry:
                        item = QTreeWidgetItem([lyr.name(), f'{lyr.featureCount():,}']); item.setData(0, Qt.UserRole, ('layer', lyr.id()))
                        (parent_item.addChild(item) if parent_item else self.groups.addTopLevelItem(item))
        walk(root); self.groups.expandAll()
        footer = QHBoxLayout(); footer.addStretch(1)
        export = QPushButton('⇩  Ekspor audit Excel…'); export.setObjectName('auditExportButton'); export.clicked.connect(self.export_csv); export.setMinimumHeight(40)
        close = QPushButton('Tutup'); close.setObjectName('auditCloseButton'); close.clicked.connect(self.close); close.setMinimumHeight(40)
        footer.addWidget(export); footer.addWidget(close); lay.addLayout(footer)
        self.setStyleSheet('''QDialog{background:#f3f6fb;font-family:'Segoe UI';font-size:10pt;color:#243247;} QLabel{color:#243247;} QComboBox{background:white;border:1px solid #cbd6e4;border-radius:7px;padding:7px;} QPushButton{background:#e8eef6;color:#203247;border:1px solid #b8c7d9;border-radius:8px;padding:8px 12px;font-weight:700;min-height:20px;} QPushButton:hover{background:#d7e8fb;border-color:#5b9bd5;} QPushButton#auditRefreshButton{background:#1769aa;color:white;border:1px solid #12568d;font-weight:800;} QPushButton#auditRefreshButton:hover{background:#0f568f;} QPushButton#auditExportButton{background:#16804a;color:white;border:1px solid #11683c;font-weight:800;} QPushButton#auditExportButton:hover{background:#10683c;} QPushButton#auditCloseButton{background:#dbe3ed;color:#26364a;border:1px solid #aab9ca;font-weight:800;} QPushButton#auditCloseButton:hover{background:#c8d4e2;} QTreeWidget{background:white;alternate-background-color:#f6f9fd;gridline-color:#e6ebf2;border:1px solid #dce4ef;border-radius:8px;selection-background-color:#cfe5ff;} QFrame#coveragePanel{background:#ffffff;border:1px solid #dce5ef;border-radius:12px;} QProgressBar{background:#e8eef6;border:0;border-radius:8px;text-align:center;color:#17324d;font-weight:800;min-height:22px;} QProgressBar::chunk{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #1677d2,stop:1 #18b879);border-radius:8px;} QHeaderView::section{background:#edf2f8;color:#42536a;padding:8px;border:0;border-bottom:1px solid #dce4ef;font-weight:700;}''')
        self.groups.itemSelectionChanged.connect(self.recount); self.reference.currentIndexChanged.connect(self.recount); self.recount()
    def selected_layer_ids(self):
        item = self.groups.currentItem()
        if not item: return []
        data = item.data(0, Qt.UserRole)
        if not data: return []
        kind, value = data
        if kind == 'layer': return [value]
        group = QgsProject.instance().layerTreeRoot().findGroup(value)
        return [n.layerId() for n in group.findLayers() if n.layer() and n.layer().isValid() and isinstance(n.layer(), QgsVectorLayer) and n.layer().geometryType() == QgsWkbTypes.PointGeometry] if group else []
    def recount(self, *args):
        ref = QgsProject.instance().mapLayer(self.reference.currentData()) if self.reference.count() else None
        self.total.setText(f'{ref.featureCount():,}' if ref else '0')
        ids = self.selected_layer_ids(); layers = [QgsProject.instance().mapLayer(i) for i in ids]; layers = [l for l in layers if l and l.isValid()]
        total_selected = sum(l.featureCount() for l in layers)
        self.result.setText(f'{total_selected:,}')
        self.layers_count.setText(str(len(layers)))
        pct = (total_selected * 100.0 / ref.featureCount()) if ref and ref.featureCount() else 0.0
        self.coverage_percent.setText(f'{pct:.1f}%')
        self.coverage_bar.setValue(max(0, min(100, int(round(pct)))))
        self.coverage_bar.setFormat(f'{pct:.1f}%  •  progres monitoring titik bangunan final')
        selection = self.groups.currentItem().text(0) if self.groups.currentItem() else 'Belum memilih grup/folder'
        self.coverage_detail.setText(f'{selection}  |  {total_selected:,} titik pekerjaan dari {ref.featureCount():,} titik acuan' if ref else 'Pilih layer acuan yang valid.')
    def export_csv(self):
        """Export the selected audit summary as a real .xlsx workbook using only stdlib."""
        path, _ = QFileDialog.getSaveFileName(self, 'Ekspor audit titik ke Excel', 'Audit_Laporan_Titik.xlsx', 'Excel Workbook (*.xlsx)')
        if not path: return
        if not path.lower().endswith('.xlsx'): path += '.xlsx'
        try:
            import zipfile, html
            ref = QgsProject.instance().mapLayer(self.reference.currentData()) if self.reference.count() else None
            ids = self.selected_layer_ids()
            layers = [QgsProject.instance().mapLayer(i) for i in ids]
            layers = [l for l in layers if l and l.isValid()]
            item = self.groups.currentItem(); selection = item.text(0) if item else '(belum dipilih)'
            total_ref = ref.featureCount() if ref else 0
            total_selected = sum(l.featureCount() for l in layers)
            rows = [
                ['Jenis', 'Nama layer/group', 'Total titik', 'Total acuan', 'Cakupan vs acuan (%)'],
                ['Acuan', ref.name() if ref else '', total_ref, total_ref, '100.00' if total_ref else '0.00'],
                ['Pilihan pekerjaan', selection, total_selected, total_ref, f'{(total_selected/total_ref*100):.2f}' if total_ref else '0.00'],
            ]
            for layer in layers:
                count = layer.featureCount()
                rows.append(['Layer pekerjaan', layer.name(), count, total_ref, f'{(count/total_ref*100):.2f}' if total_ref else '0.00'])
            def col_name(n):
                out = ''
                while n:
                    n, r = divmod(n-1, 26); out = chr(65+r) + out
                return out
            sheet_rows = []
            for ri, row in enumerate(rows, 1):
                cells = []
                for ci, value in enumerate(row, 1):
                    refcell = f'{col_name(ci)}{ri}'
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        cells.append(f'<c r="{refcell}" t="n"><v>{value}</v></c>')
                    else:
                        val = html.escape(str(value), quote=False)
                        cells.append(f'<c r="{refcell}" t="inlineStr"><is><t xml:space="preserve">{val}</t></is></c>')
                sheet_rows.append(f'<row r="{ri}">{"".join(cells)}</row>')
            sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + ''.join(sheet_rows) + '</sheetData></worksheet>'
            workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Audit Titik" sheetId="1" r:id="rId1"/></sheets></workbook>'
            rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
            root_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
            types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
            with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', types)
                z.writestr('_rels/.rels', root_rels)
                z.writestr('xl/workbook.xml', workbook)
                z.writestr('xl/_rels/workbook.xml.rels', rels)
                z.writestr('xl/worksheets/sheet1.xml', sheet)
            QMessageBox.information(self, 'Ekspor selesai', 'Laporan audit Excel (.xlsx) berhasil dibuat.\n' + path)
        except Exception as e:
            QMessageBox.critical(self, 'Ekspor gagal', str(e))

class AutoSnapBangunanPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.action = None
        self.style_action = None
        self.village_style_action = None
        self.audit_action = None
        self.dialog = None
        self.main_menu = None
        self.geotag_menu = None
        self.toolbox_menu = None
        self.manual_action = None
        self.automatic_action = None

    def initGui(self):
        # v2.6.5: visible QGIS menu + popup command groups. Processing methods are unchanged.
        icon_path = os.path.join(os.path.dirname(__file__), 'icon.svg')
        main_window = self.iface.mainWindow()
        self.action = QAction(QIcon(icon_path), 'SLS Magic AutoSnap Bangunan', main_window)
        self.action.setToolTip('Buka panel utama AutoSnap')
        self.action.triggered.connect(lambda checked=False: self.run())
        self.iface.addToolBarIcon(self.action)

        self.main_menu = QMenu('SLS MAGIC', main_window)
        self.main_menu.setToolTipsVisible(True)
        self.main_menu.addAction(self.action)
        self.main_menu.addSeparator()

        self.geotag_menu = QMenu('Auto Geotagging', self.main_menu)
        self.manual_action = QAction('Manual', main_window)
        self.manual_action.triggered.connect(lambda checked=False: self.run(0))
        self.geotag_menu.addAction(self.manual_action)
        self.automatic_action = QAction('Otomatis', main_window)
        self.automatic_action.triggered.connect(lambda checked=False: self.run(1))
        self.geotag_menu.addAction(self.automatic_action)
        self.main_menu.addMenu(self.geotag_menu)

        self.toolbox_menu = QMenu('Toolbox', self.main_menu)
        self.village_style_action = QAction('Style Layer', main_window)
        self.village_style_action.triggered.connect(self.style_by_village)
        self.toolbox_menu.addAction(self.village_style_action)
        self.style_action = QAction('Load Style — Titik Bangunan', main_window)
        self.style_action.triggered.connect(self.load_style)
        self.toolbox_menu.addAction(self.style_action)
        self.main_menu.addMenu(self.toolbox_menu)

        self.main_menu.addSeparator()
        self.audit_action = QAction('Audit Laporan Titik', main_window)
        self.audit_action.triggered.connect(self.open_audit)
        self.main_menu.addAction(self.audit_action)
        self.update_action = QAction('Cek Update GitHub', main_window)
        self.update_action.setToolTip('Periksa rilis terbaru SLS Magic AutoSnap Bangunan di GitHub')
        self.update_action.triggered.connect(self.check_github_update)
        self.main_menu.addAction(self.update_action)
        # Attach visibly to the QGIS menu bar, not as a nested submenu under Plugins.
        self.iface.mainWindow().menuBar().addMenu(self.main_menu)

    def check_github_update(self):
        """Check latest published release and offer a safe staged install."""
        api = 'https://api.github.com/repos/fiki-hub/SLS-Magic-AutoSnap-Bangunan/releases/latest'
        releases_url = 'https://github.com/fiki-hub/SLS-Magic-AutoSnap-Bangunan/releases'
        try:
            req = urllib.request.Request(api, headers={'User-Agent': 'SLS-Magic-AutoSnap-Bangunan', 'Accept': 'application/vnd.github+json'})
            with urllib.request.urlopen(req, timeout=12) as response:
                release = json.loads(response.read().decode('utf-8'))
            tag = str(release.get('tag_name') or '')
            assets = release.get('assets') or []
            zip_asset = next((a for a in assets if str(a.get('name','')).lower().endswith('.zip')), None)
            if not tag:
                QMessageBox.information(self.iface.mainWindow(), 'Update GitHub',
                    'Belum ada GitHub Release yang diterbitkan. Source repository tersedia, tetapi update otomatis memerlukan ZIP pada GitHub Releases.\n\nBuka halaman Releases untuk melihat statusnya.')
                webbrowser.open(releases_url)
                return
            current = 'unknown'
            try:
                with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'metadata.txt'), 'r', encoding='utf-8') as meta_file:
                    for meta_line in meta_file:
                        if meta_line.lower().startswith('version='):
                            current = meta_line.split('=', 1)[1].strip()
                            break
            except Exception:
                pass
            remote = tag.lstrip('vV')
            if remote == current:
                QMessageBox.information(self.iface.mainWindow(), 'Update GitHub', 'Plugin sudah menggunakan versi {}.'.format(current))
                return
            if not zip_asset or not zip_asset.get('browser_download_url'):
                QMessageBox.warning(self.iface.mainWindow(), 'Update GitHub', 'Rilis {} belum memiliki aset ZIP. Tidak ada file yang diubah. Unggah ZIP plugin ke GitHub Releases terlebih dahulu.'.format(tag))
                webbrowser.open(releases_url)
                return
            answer = QMessageBox.question(self.iface.mainWindow(), 'Update GitHub',
                'Versi rilis tersedia: {}\nVersi saat ini: {}\n\nUnduh dan siapkan update? QGIS perlu ditutup lalu dibuka kembali agar modul baru dimuat.'.format(tag, current),
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if answer != QMessageBox.Yes:
                return
            plugin_dir = os.path.dirname(os.path.abspath(__file__))
            parent_dir = os.path.dirname(plugin_dir)
            stage = tempfile.mkdtemp(prefix='slsmagic_update_', dir=parent_dir)
            archive_path = os.path.join(stage, 'release.zip')
            req = urllib.request.Request(zip_asset['browser_download_url'], headers={'User-Agent': 'SLS-Magic-AutoSnap-Bangunan'})
            with urllib.request.urlopen(req, timeout=45) as response, open(archive_path, 'wb') as out:
                shutil.copyfileobj(response, out)
            import zipfile
            with zipfile.ZipFile(archive_path, 'r') as archive:
                names = archive.namelist()
                roots = set(n.split('/')[0] for n in names if '/' in n)
                if not names or len(roots) != 1:
                    raise ValueError('Struktur ZIP tidak sesuai: harus berisi satu folder plugin utama.')
                root_name = next(iter(roots))
                archive.extractall(stage)
            extracted = os.path.join(stage, root_name)
            if not os.path.isfile(os.path.join(extracted, 'metadata.txt')) or not os.path.isfile(os.path.join(extracted, '__init__.py')):
                raise ValueError('ZIP tidak memuat metadata.txt dan __init__.py di folder utama.')
            pending = plugin_dir + '_update_pending'
            backup = plugin_dir + '_backup'
            if os.path.exists(pending):
                shutil.rmtree(pending, ignore_errors=True)
            shutil.copytree(extracted, pending, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
            # Do not replace loaded Python modules while QGIS is running. Stage the new
            # package and a helper BAT; user can close QGIS and run it to swap directories.
            bat_path = os.path.join(stage, 'apply_update.bat')
            with open(bat_path, 'w', encoding='utf-8') as bat:
                bat.write('@echo off\r\n')
                bat.write('echo Close QGIS before applying this update.\r\n')
                bat.write('if exist "{}" rmdir /s /q "{}"\r\n'.format(backup, backup))
                bat.write('if exist "{}" move "{}" "{}"\r\n'.format(plugin_dir, plugin_dir, backup))
                bat.write('move "{}" "{}"\r\n'.format(pending, plugin_dir))
                bat.write('echo Update staged. Reopen QGIS and verify the plugin.\r\n')
                bat.write('pause\r\n')
            QMessageBox.information(self.iface.mainWindow(), 'Update disiapkan',
                'Paket {} sudah diunduh dan divalidasi.\n\n1. Tutup QGIS sepenuhnya.\n2. Jalankan apply_update.bat di folder: {}\n3. Buka QGIS kembali.\n\nFolder cadangan dibuat sebagai {}.'.format(tag, stage, backup))
            os.startfile(stage)
        except Exception as exc:
            QMessageBox.warning(self.iface.mainWindow(), 'Update GitHub',
                'Update tidak diterapkan dan file aktif tidak diganti. Detail: {}'.format(exc))
            webbrowser.open(releases_url)

    def open_audit(self):
        dlg = PointAuditDialog(self.iface, self.iface.mainWindow())
        dlg.exec_()
    def unload(self):
        if self.dialog:
            try:
                self.dialog.stop_pick(restore=True)
            except Exception:
                pass
            try:
                self.dialog.restore_source_subset()
            except Exception:
                pass
        if self.action is not None:
            self.iface.removeToolBarIcon(self.action)
        if self.main_menu is not None:
            try:
                self.iface.mainWindow().menuBar().removeAction(self.main_menu.menuAction())
            except Exception:
                pass
            self.main_menu.deleteLater()
        self.main_menu = None
        self.geotag_menu = None
        self.toolbox_menu = None

    def run(self, tab_index=None):
        # Reuse the same dialog instance; selecting a menu item switches its existing tab.
        import traceback
        try:
            if self.dialog is None:
                self.dialog = AutoSnapDialog(self.iface, self.iface.mainWindow())
            self.dialog.refresh_layers(force=False)
            if tab_index is not None and hasattr(self.dialog, 'tabs'):
                self.dialog.tabs.setCurrentIndex(max(0, min(int(tab_index), self.dialog.tabs.count() - 1)))
            self.dialog.show()
            self.dialog.raise_()
            self.dialog.activateWindow()
        except Exception as e:
            self.dialog = None
            QMessageBox.critical(self.iface.mainWindow(), 'SLS Magic AutoSnap gagal dibuka',
                                 f'{e}\n\n{traceback.format_exc()}')

    def style_by_village(self):
        layer=self.iface.activeLayer()
        if not isinstance(layer, QgsVectorLayer):
            QMessageBox.warning(self.iface.mainWindow(), 'Pilih layer', 'Aktifkan layer georeferencing/titik terlebih dahulu.'); return
        fields=[f.name() for f in layer.fields()]
        preferred=[f for f in fields if any(k in f.lower() for k in ('desa','village','kelurahan'))]
        choices=preferred or [f for f in fields if any(k in f.lower() for k in ('sls','level_6_name'))]
        if not choices:
            QMessageBox.warning(self.iface.mainWindow(), 'Kolom desa tidak ditemukan', 'Layer aktif tidak memiliki kolom nama desa. Tambahkan/isi atribut desa terlebih dahulu. Kolom level_6_name dapat dipakai untuk warna per SLS.'); return
        field, ok=QInputDialog.getItem(self.iface.mainWindow(), 'Style per wilayah', 'Kelompokkan warna berdasarkan kolom:', choices, 0, False)
        if not ok or not field: return
        values=sorted({str(f[field]).strip() if f[field] not in (None, '') else '(kosong)' for f in layer.getFeatures()}, key=str.lower)
        palette=['#2786d9','#20a486','#ef8b35','#9561d8','#d94f70','#26a6b8','#8b9b2b','#e0b323','#5166d6','#c05c35','#458b72','#a65fb5']
        cats=[]
        for i,val in enumerate(values):
            col=QColor(palette[i % len(palette)])
            sym=QgsMarkerSymbol.createSimple({'name':'circle','color':f'{col.red()},{col.green()},{col.blue()},255','outline_color':'35,45,60,255','outline_width':'0.25','size':'4'}) if layer.geometryType()==QgsWkbTypes.PointGeometry else None
            if sym is None:
                QMessageBox.information(self.iface.mainWindow(), 'Jenis geometri', 'Layer bukan titik. Untuk poligon, gunakan symbologi kategori QGIS pada kolom desa.'); return
            cats.append(QgsRendererCategory(val, sym, val))
        layer.setRenderer(QgsCategorizedSymbolRenderer(field, cats))
        # Keep labels for point building layers that have no_bang.
        if 'no_bang' in fields:
            temp=AutoSnapDialog.__new__(AutoSnapDialog); temp.source=layer
            AutoSnapDialog.apply_style(temp, layer, QColor('#2786d9'))
            # re-apply categorized renderer after shared label style helper
            layer.setRenderer(QgsCategorizedSymbolRenderer(field, cats))
        layer.triggerRepaint(); self.iface.mapCanvas().refresh()
        self.iface.messageBar().pushSuccess('SLS MAGIC', f'Style kategori diterapkan: {len(values)} nilai pada kolom {field}.')

    def load_style(self):
        layer=self.iface.activeLayer()
        if not isinstance(layer,QgsVectorLayer) or layer.geometryType()!=QgsWkbTypes.PointGeometry or 'no_bang' not in [f.name() for f in layer.fields()]:
            QMessageBox.warning(self.iface.mainWindow(),'Pilih layer titik','Aktifkan layer titik bangunan yang memiliki field no_bang, lalu jalankan Load Style.'); return
        dlg=QMessageBox(self.iface.mainWindow()); dlg.setWindowTitle('Load Style'); dlg.setText(f'Terapkan symbologi dan label pada layer “{layer.name()}”?'); dlg.setInformativeText('Ukuran titik 4, satu warna hijau seragam untuk seluruh layer (bukan warna per FID), label no_bang putih tebal dengan buffer/bayangan hitam. Layer output AutoSnap memakai satu warna berbeda per layer.'); dlg.setStandardButtons(QMessageBox.Yes|QMessageBox.No)
        if dlg.exec_()==QMessageBox.Yes:
            temp=AutoSnapDialog.__new__(AutoSnapDialog); temp.source=layer
            AutoSnapDialog.apply_style(temp,layer)
            self.iface.messageBar().pushSuccess('SLS MAGIC','Style diterapkan ke layer aktif.')
