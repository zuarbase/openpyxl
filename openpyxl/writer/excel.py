# Copyright (c) 2010-2022 openpyxl

"""Write a .xlsx file."""

# Python stdlib imports
import re
from tempfile import TemporaryFile
from zipfile import ZipFile, ZIP_DEFLATED

# package imports
from openpyxl.compat import deprecated
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.xml.constants import (
    ARC_ROOT_RELS,
    ARC_WORKBOOK_RELS,
    ARC_APP, ARC_CORE,
    ARC_CUSTOM, CPROPS_TYPE,
    ARC_THEME,
    ARC_STYLE,
    ARC_WORKBOOK,
    IMAGE_NS,
)
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.drawing.legacy import LegacyDrawing
from openpyxl.xml.functions import tostring
from openpyxl.packaging.manifest import Manifest
from openpyxl.packaging.relationship import (
    get_rels_path,
    RelationshipList,
    Relationship,
)
from openpyxl.comments.comment_sheet import CommentSheet
from openpyxl.packaging.extended import ExtendedProperties
from openpyxl.styles.stylesheet import write_stylesheet
from openpyxl.worksheet._writer import WorksheetWriter
from openpyxl.workbook._writer import WorkbookWriter
from .theme import theme_xml


class ExcelWriter(object):
    """Write a workbook object to an Excel file."""


    def __init__(self, workbook, archive):
        self.archive = archive
        self.workbook = workbook
        self.manifest = Manifest()
        self.vba_modified = set()
        self._tables = []
        self._charts = []
        self._images = []
        self._drawings = []
        self._comments = []
        self._pivots = []
        self.activex = []
        self.legacy = []
        self.form_controls = []


    def write_data(self):
        """Write the various xml files into the zip archive."""
        # cleanup all worksheets
        archive = self.archive

        props = ExtendedProperties()
        archive.writestr(ARC_APP, tostring(props.to_tree()))

        archive.writestr(ARC_CORE, tostring(self.workbook.properties.to_tree()))
        if self.workbook.loaded_theme:
            archive.writestr(ARC_THEME, self.workbook.loaded_theme)
        else:
            archive.writestr(ARC_THEME, theme_xml)

        if len(self.workbook.custom_doc_props) >= 1:
            archive.writestr(ARC_CUSTOM, tostring(self.workbook.custom_doc_props.to_tree()))
            class CustomOverride():
                path = "/" + ARC_CUSTOM #PartName
                mime_type = CPROPS_TYPE #ContentType

            custom_override = CustomOverride()
            # custom_override = Override(PartName="/docProps/custom.xml", ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml")
            self.manifest.append(custom_override)

        self._write_worksheets()
        self._write_chartsheets()
        self._write_images()
        self._write_charts()

        self._write_external_links()

        stylesheet = write_stylesheet(self.workbook)
        archive.writestr(ARC_STYLE, tostring(stylesheet))

        writer = WorkbookWriter(self.workbook)
        archive.writestr(ARC_ROOT_RELS, writer.write_root_rels())
        archive.writestr(ARC_WORKBOOK, writer.write())
        archive.writestr(ARC_WORKBOOK_RELS, writer.write_rels())

        self._merge_vba()

        self.manifest._write(archive, self.workbook)


    def _add_image(self, img):
        """
        Collect images from various components
        """
        if img in self._images:
            img._id = self._images.index(img)
        else:
            self._images.append(img)
            img._id = len(self._images)


    def _merge_vba(self):
        """
        If workbook contains macros then extract associated files from cache
        of old file and add to archive
        """
        if self.workbook._vba:
            self.archive.writestr("xl/vbaProject.bin", self.workbook._vba)
        #ARC_VBA = re.compile("|".join(
            #('xl/vba','customUI', ))
                             #)

        #if self.workbook.vba_archive:
            #for name in set(self.workbook.vba_archive.namelist()) - self.vba_modified:
                #if ARC_VBA.match(name):
                    #self.archive.writestr(name, self.workbook.vba_archive.read(name))


    def _write_images(self):
        # delegate to object
        for img in self._images:
            img._write(self.archive)


    def _write_charts(self):
        # delegate to object
        if len(self._charts) != len(set(self._charts)):
            raise InvalidFileException("The same chart cannot be used in more than one worksheet")
        for chart in self._charts:
            self.archive.writestr(chart.path[1:], tostring(chart._write()))
            self.manifest.append(chart)


    def _write_drawing(self, drawing):
        """
        Write a drawing
        """
        self._drawings.append(drawing)
        drawing._id = len(self._drawings)
        for chart in drawing.charts:
            self._charts.append(chart)
            chart._id = len(self._charts)
        for img in drawing.images:
            self._images.append(img)
            img._id = len(self._images)
        rels_path = get_rels_path(drawing.path)[1:]
        self.archive.writestr(drawing.path[1:], tostring(drawing._write()))
        self.archive.writestr(rels_path, tostring(drawing._write_rels()))
        self.manifest.append(drawing)


    def _write_chartsheets(self):
        for idx, sheet in enumerate(self.workbook.chartsheets, 1):

            sheet._id = idx
            xml = tostring(sheet.to_tree())

            self.archive.writestr(sheet.path[1:], xml)
            self.manifest.append(sheet)

            if sheet._drawing:
                self._write_drawing(sheet._drawing)

                rel = Relationship(type="drawing", Target=sheet._drawing.path)
                rels = RelationshipList()
                rels.append(rel)
                tree = rels.to_tree()

                rels_path = get_rels_path(sheet.path[1:])
                self.archive.writestr(rels_path, tostring(tree))


    def _write_comment(self, ws):

        cs = CommentSheet.from_comments(ws._comments)
        self._comments.append(cs)
        cs._id = len(self._comments)
        self.archive.writestr(cs.path[1:], tostring(cs.to_tree()))
        self.manifest.append(cs)

        vml = None
        if ws.legacy_drawing is not None:
            vml = ws.legacy_drawing.vml
        else:
            ws.legacy_drawing = LegacyDrawing(vml)

        vml = cs.write_shapes(vml)
        ws.legacy_drawing.vml = vml

        comment_rel = Relationship(Id="comments", type=cs._rel_type, Target=cs.path)
        ws._rels.append(comment_rel)


    def write_legacy(self, ws):
        """
        Write VML
        """
        drawing = ws.legacy_drawing
        if ws.legacy_drawing is None:
            return

        self.legacy.append(drawing)
        drawing._counter = len(self.legacy)
        drawing._write(self.archive)

        for rel in drawing.children.Relationship:
            img = rel.blob
            self._add_image(img)
            rel.Target = img.path

        if drawing.children:
            xml = drawing.children.to_tree()
            path = get_rels_path(ws.legacy_drawing.path)
            self.archive.writestr(path[1:], tostring(xml))


    def write_worksheet(self, ws):
        ws._drawing = SpreadsheetDrawing()
        ws._drawing.charts = ws._charts
        ws._drawing.images = ws._images
        ws._drawing.shapes = ws._shapes

        if self.workbook.write_only:
            if not ws.closed:
                ws.close()
            writer = ws._writer
        else:
            writer = WorksheetWriter(ws)
            writer.write()

        ws._rels = writer._rels

        if ws._comments:
            self._write_comment(ws)
        self.write_controls(ws, writer.controls)
        self.write_embedded(ws, writer.control_images)

        if ws.legacy_drawing:
            drawing = ws.legacy_drawing
            self.write_legacy(ws)
            rel = ws._rels[drawing._rel_id]
            rel.Target = drawing.path

        self.archive.write(writer.out, ws.path[1:])
        self.manifest.append(ws)

        writer.cleanup()


    def write_controls(self, ws, controls):
        """Serialise form controls for a specific worksheet"""

        for ctrl in controls:
            if hasattr(ctrl, "bin"): # ugh!
                store = self.activex
            else:
                store = self.form_controls
            store.append(ctrl)
            ctrl.counter = len(store) # ugh!
            ctrl._write(self.archive, self.manifest)

            ws._rels[ctrl._rel_id].Target = ctrl.path


    def write_embedded(self, ws, control_images):
        """
        Update path for embedded images for controls
        """
        for obj in control_images:
            img = obj.blob
            self._add_image(img)
            obj.Target = img.path


    def _write_worksheets(self):

        pivot_caches = set()

        for idx, ws in enumerate(self.workbook.worksheets, 1):
            ws._id = idx
            self.write_worksheet(ws)

            if ws._drawing:
                self._write_drawing(ws._drawing)
                for r in ws._rels.Relationship:
                    if "drawing" in r.Type:
                        r.Target = ws._drawing.path

            for t in ws._tables.values():
                self._tables.append(t)
                t.id = len(self._tables)
                t._write(self.archive)
                self.manifest.append(t)
                ws._rels[t._rel_id].Target = t.path

            for p in ws._pivots:
                if p.cache not in pivot_caches:
                    pivot_caches.add(p.cache)
                    p.cache._id = len(pivot_caches)

                self._pivots.append(p)
                p._id = len(self._pivots)
                p._write(self.archive, self.manifest)
                self.workbook._pivots.append(p)
                r = Relationship(Type=p.rel_type, Target=p.path)
                ws._rels.append(r)

            if ws._rels:
                tree = ws._rels.to_tree()
                rels_path = get_rels_path(ws.path)[1:]
                self.archive.writestr(rels_path, tostring(tree))


    def _write_external_links(self):
        # delegate to object
        """Write links to external workbooks"""
        wb = self.workbook
        for idx, link in enumerate(wb._external_links, 1):
            link._id = idx
            rels_path = get_rels_path(link.path[1:])

            xml = link.to_tree()
            self.archive.writestr(link.path[1:], tostring(xml))
            rels = RelationshipList()
            rels.append(link.file_link)
            self.archive.writestr(rels_path, tostring(rels.to_tree()))
            self.manifest.append(link)


    def save(self):
        """Write data into the archive."""
        self.write_data()
        self.archive.close()


def save_workbook(workbook, filename):
    """Save the given workbook on the filesystem under the name filename.

    :param workbook: the workbook to save
    :type workbook: :class:`openpyxl.workbook.Workbook`

    :param filename: the path to which save the workbook
    :type filename: string

    :rtype: bool

    """
    #if wb._vba and not filename.endswith(".xlsm"):
        #warn()
    archive = ZipFile(filename, 'w', ZIP_DEFLATED, allowZip64=True)
    writer = ExcelWriter(workbook, archive)
    writer.save()
    return True


@deprecated("Use a NamedTemporaryFile")
def save_virtual_workbook(workbook):
    """Return an in-memory workbook, suitable for a Django response."""
    tmp = TemporaryFile()
    archive = ZipFile(tmp, 'w', ZIP_DEFLATED, allowZip64=True)

    writer = ExcelWriter(workbook, archive)
    writer.save()

    tmp.seek(0)
    virtual_workbook = tmp.read()
    tmp.close()

    return virtual_workbook
