# Copyright (c) 2010-2022 openpyxl
import pytest

from io import BytesIO
from zipfile import ZipFile

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from openpyxl.worksheet.filters import (
    AutoFilter,
    Filters,
    FilterColumn,
)


@pytest.fixture
def TableColumn():
    from ..table import TableColumn
    return TableColumn


class TestTableColumn:

    def test_ctor(self, TableColumn):
        col = TableColumn(id=1, name="Column1")
        xml = tostring(col.to_tree())
        expected = """
        <tableColumn id="1" name="Column1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TableColumn):
        src = """
        <tableColumn id="1" name="Column1"/>
        """
        node = fromstring(src)
        col = TableColumn.from_tree(node)
        assert col == TableColumn(id=1, name="Column1")


@pytest.fixture
def Table():
    from ..table import Table
    return Table


class TestTable:

    def test_ctor(self, Table, TableColumn):
        table = Table(displayName="A_Sample_Table", ref="A1:D5")
        xml = tostring(table.to_tree())
        expected = """
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           displayName="A_Sample_Table" headerRowCount="1" name="A_Sample_Table" id="1" ref="A1:D5">
        </table>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_columns(self, Table):
        table = Table(displayName="A_Sample_Table", ref="A1:D5")
        table._initialise_columns()
        xml = tostring(table.to_tree())
        expected = """
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           displayName="A_Sample_Table" headerRowCount="1" name="A_Sample_Table" id="1" ref="A1:D5">
           <autoFilter ref="A1:D5" />
        <tableColumns count="4">
          <tableColumn id="1" name="Column1" />
          <tableColumn id="2" name="Column2" />
          <tableColumn id="3" name="Column3" />
          <tableColumn id="4" name="Column4" />
        </tableColumns>
        </table>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_preserve_existing_filter(self, Table):
        table = Table(displayName="A_Sample_Table", ref="A1:D5")
        filters = Filters(blank=False, filter=["16"])
        col = FilterColumn(colId=0, filters=filters)
        table.autoFilter = AutoFilter(ref="A1:D5", filterColumn=[col])
        table._initialise_columns()
        assert table.autoFilter.filterColumn == [col]


    def test_column_names(self, Table):
        table = Table(displayName="A_Sample_Table", ref="A10:D14")
        table._initialise_columns()
        assert table.column_names == ["Column1", "Column2", "Column3", "Column4"]


    def test_from_xml(self, Table):
        src = """
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            id="1" name="Table1" displayName="Table1" ref="A1:AA27">
        </table>
        """
        node = fromstring(src)
        table = Table.from_tree(node)
        assert table == Table(displayName="Table1", name="Table1",
                              ref="A1:AA27")


    def test_path(self, Table):
        table = Table(displayName="Table1", ref="A1:M6")
        assert table.path == "/xl/tables/table1.xml"


    def test_write(self, Table):
        out = BytesIO()
        archive = ZipFile(out, "w")
        table = Table(displayName="Table1", ref="B1:L10")
        table._write(archive)
        assert "xl/tables/table1.xml" in archive.namelist()


@pytest.fixture
def TableFormula():
    from ..table import TableFormula
    return TableFormula


class TestTableFormula:

    def test_ctor(self, TableFormula):
        formula = TableFormula()
        formula.text = "=A1*4"
        xml = tostring(formula.to_tree())
        expected = """
        <tableFormula>=A1*4</tableFormula>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TableFormula):
        src = """
        <tableFormula>=A1*4</tableFormula>
        """
        node = fromstring(src)
        formula = TableFormula.from_tree(node)
        assert formula.text == "=A1*4"


@pytest.fixture
def TableStyleInfo():
    from ..table import TableStyleInfo
    return TableStyleInfo


class TestTableInfo:

    def test_ctor(self, TableStyleInfo):
        info = TableStyleInfo(name="TableStyleMedium12")
        xml = tostring(info.to_tree())
        expected = """
        <tableStyleInfo name="TableStyleMedium12" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TableStyleInfo):
        src = """
        <tableStyleInfo name="TableStyleLight1" showRowStripes="1" />
        """
        node = fromstring(src)
        info = TableStyleInfo.from_tree(node)
        assert info == TableStyleInfo(name="TableStyleLight1", showRowStripes=True)


@pytest.fixture
def XMLColumnProps():
    from ..table import XMLColumnProps
    return XMLColumnProps


class TestXMLColumnPr:

    def test_ctor(self, XMLColumnProps):
        col = XMLColumnProps(mapId="1", xpath="/xml/foo/element", xmlDataType="string")
        xml = tostring(col.to_tree())
        expected = """
        <xmlColumnPr mapId="1" xpath="/xml/foo/element" xmlDataType="string"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, XMLColumnProps):
        src = """
        <xmlColumnPr mapId="1" xpath="/xml/foo/element" xmlDataType="string"/>
        """
        node = fromstring(src)
        col = XMLColumnProps.from_tree(node)
        assert col == XMLColumnProps(mapId="1", xpath="/xml/foo/element", xmlDataType="string")



@pytest.fixture
def TablePartList():
    from ..table import TablePartList
    return TablePartList

from ..related import Related


class TestTablePartList:

    def test_ctor(self, TablePartList):
        tables = TablePartList()
        tables.append(Related(id="rId1"))
        xml = tostring(tables.to_tree())
        expected = """
        <tableParts count="1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <tablePart r:id="rId1" />
        </tableParts>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TablePartList):
        src = """
        <tableParts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <tablePart r:id="rId1" />
          <tablePart r:id="rId2" />
        </tableParts>
        """
        node = fromstring(src)
        tables = TablePartList.from_tree(node)
        assert len(tables.tablePart) == 2


@pytest.fixture
def TableList():
    from ..table import TableList
    return TableList


class TestTableList:

    def test_append(self, Table, TableList):
        tablelist = TableList()
        table1 = Table(displayName="Table1", ref="A1:C10")
        tablelist.add(table1)
        assert len(tablelist) == 1


    def test_get(self, Table, TableList):
        tablelist = TableList()
        table1 = Table(displayName="Table1", ref="A1:C10")
        tablelist.add(table1)
        assert table1 == tablelist["Table1"]


    def test_get_by_range(self, Table, TableList):
        tablelist = TableList()
        table1 = Table(displayName="Table1", ref="A1:D10")
        tablelist.add(table1)
        assert True == isinstance(tablelist.get(table_range="A1:D10"),Table)


    def test_add_type_error(self, Table, TableList):
        tablelist = TableList()
        with pytest.raises(TypeError):
            tablelist.add("Not a Table")


    def test_get_table_does_not_exists(self, Table, TableList):
        tablelist2 = TableList()
        with pytest.raises(KeyError):
            tablelist2['NoTable']


    def test_items(self, Table, TableList):
        table1 = Table(displayName="Table1", ref="A1:D10")
        table2 = Table(displayName="Table2", ref="A1:D10")
        tablelist = TableList()
        tablelist.add(table1)
        tablelist.add(table2)
        assert tablelist.items() == [("Table1", "A1:D10"), ("Table2", "A1:D10")]
