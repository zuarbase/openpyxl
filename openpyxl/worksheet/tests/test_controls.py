# Copyright (c) 2010-2021 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from openpyxl.worksheet.ole import ObjectAnchor, AnchorMarker


@pytest.fixture
def ControlProperty():
    from ..controls import ControlProperty
    return ControlProperty


class TestControlProperty:

    def test_ctor(self, ControlProperty):
        _from = AnchorMarker()
        to = AnchorMarker()
        anchor = ObjectAnchor(_from=_from, to=to)
        prop = ControlProperty(anchor=anchor)
        xml = tostring(prop.to_tree())
        expected = """
        <controlPr autoFill="1" autoLine="1" autoPict="1" cf="pict" defaultSize="1" disabled="0" locked="1" print="1"
        recalcAlways="0" uiObject="0"
        xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
          <anchor moveWithCells="0" sizeWithCells="0">
            <from>
              <xdr:col>0</xdr:col>
              <xdr:colOff>0</xdr:colOff>
              <xdr:row>0</xdr:row>
              <xdr:rowOff>0</xdr:rowOff>
            </from>
            <to>
              <xdr:col>0</xdr:col>
              <xdr:colOff>0</xdr:colOff>
              <xdr:row>0</xdr:row>
              <xdr:rowOff>0</xdr:rowOff>
            </to>
          </anchor>
        </controlPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ControlProperty):
        src = """
        <controlPr
        xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        autoLine="0">
        <anchor moveWithCells="1">
          <from>
            <xdr:col>4</xdr:col>
            <xdr:colOff>704850</xdr:colOff>
            <xdr:row>59</xdr:row>
            <xdr:rowOff>114300</xdr:rowOff>
          </from>
          <to>
            <xdr:col>4</xdr:col>
            <xdr:colOff>1190625</xdr:colOff>
            <xdr:row>61</xdr:row>
            <xdr:rowOff>47625</xdr:rowOff>
          </to>
        </anchor>
        </controlPr>
        """
        node = fromstring(src)
        prop = ControlProperty.from_tree(node)
        _from = AnchorMarker(col=4, colOff=704850, row=59, rowOff=114300)
        to = AnchorMarker(col=4, colOff=1190625, row=61, rowOff=47625)
        anchor = ObjectAnchor(_from=_from, to=to, moveWithCells=True)
        assert prop == ControlProperty(anchor=anchor, autoLine=False)


@pytest.fixture
def Control():
    from ..controls import Control
    return Control


class TestControl:

    def test_ctor(self, Control):
        ctrl = Control(shapeId=1)
        xml = tostring(ctrl.to_tree())
        expected = """
        <control shapeId="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Control):
        src = """
         <control xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" shapeId="47120" r:id="rId8" name="RefProjectButton" />
        """
        node = fromstring(src)
        ctrl = Control.from_tree(node)
        assert ctrl == Control(shapeId=47120, name="RefProjectButton", id="rId8")


@pytest.fixture
def ControlList():
    from ..controls import ControlList
    return ControlList


class TestControlList:

    def test_ctor(self, ControlList):
        ctrls = ControlList()
        xml = tostring(ctrls.to_tree())
        expected = """
        <controls />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ControlList):
        src = """
          <controls xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
              <mc:Choice Requires="x14">
              <control shapeId="47120" r:id="rId8" name="RefWECProjectButton">
                <controlPr defaultSize="0" autoLine="0" r:id="rId9">
                  <anchor moveWithCells="1">
                    <from>
                      <xdr:col>3</xdr:col>
                      <xdr:colOff>161925</xdr:colOff>
                      <xdr:row>24</xdr:row>
                      <xdr:rowOff>95250</xdr:rowOff>
                    </from>
                    <to>
                      <xdr:col>6</xdr:col>
                      <xdr:colOff>171450</xdr:colOff>
                      <xdr:row>27</xdr:row>
                      <xdr:rowOff>95250</xdr:rowOff>
                    </to>
                  </anchor>
                </controlPr>
              </control>
            </mc:Choice>
            <mc:Fallback>
              <control shapeId="47120" r:id="rId8" name="RefWECProjectButton"/>
            </mc:Fallback>
          </mc:AlternateContent>
        </controls>
        """
        node = fromstring(src)
        ctrls = ControlList.from_tree(node)
        assert len(ctrls) == 1


@pytest.fixture
def FormControl():
    from ..controls import FormControl
    return FormControl


class TestFormControl:

    def test_ctor(self, FormControl):
        ctrl = FormControl(objectType="Button", lockText=True)
        xml = tostring(ctrl.to_tree())
        expected = """
        <formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Button" lockText="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, FormControl):
        src = """
        <formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Button" lockText="1"/>
        """
        node = fromstring(src)
        ctrl = FormControl.from_tree(node)
        assert ctrl == FormControl(objectType="Button", lockText=True)
