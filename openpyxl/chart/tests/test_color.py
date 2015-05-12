from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def ColorChoice():
    from ..colors import ColorChoice
    return ColorChoice


class TestColorChoice:

    def test_ctor(self, ColorChoice):
        color = ColorChoice()
        color.RGB = "000000"
        xml = tostring(color.to_tree())
        expected = """
        <colorChoice xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
          <srgbClr val="000000" />
        </colorChoice>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ColorChoice):
        src = """
        <colorChoice />
        """
        node = fromstring(src)
        color = ColorChoice.from_tree(node)
        assert color == ColorChoice()


@pytest.fixture
def SystemColor():
    from ..colors import SystemColor
    return SystemColor


class TestSystemColor:

    def test_ctor(self, SystemColor):
        colors = SystemColor()
        xml = tostring(colors.to_tree())
        expected = """
        <sysClr val="bg1"></sysClr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SystemColor):
        src = """
        <sysClr val="tx1"></sysClr>
        """
        node = fromstring(src)
        colors = SystemColor.from_tree(node)
        assert colors == SystemColor(val="tx1")


@pytest.fixture
def HSLColor():
    from ..colors import HSLColor
    return HSLColor


class TestHSLColor:

    def test_ctor(self, HSLColor):
        colors = HSLColor(hue=50, sat=10, lum=90)
        xml = tostring(colors.to_tree())
        expected = """
        <hslClr hue="50" lum="90" sat="10" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, HSLColor):
        src = """
        <hslClr hue="0" lum="70" sat="20" />
        """
        node = fromstring(src)
        colors = HSLColor.from_tree(node)
        assert colors == HSLColor(hue=0, sat=20, lum=70)


@pytest.fixture
def RGBPercent():
    from ..colors import RGBPercent
    return RGBPercent


class TestRGBPercent:

    def test_ctor(self, RGBPercent):
        colors = RGBPercent(r=30, g=40, b=20)
        xml = tostring(colors.to_tree())
        expected = """
        <rgbClr b="20" g="40" r="30" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RGBPercent):
        src = """
        <rgbClr b="20" g="40" r="30" />
        """
        node = fromstring(src)
        colors = RGBPercent.from_tree(node)
        assert colors == RGBPercent(r=30, g=40, b=20)


@pytest.fixture
def ColorMapping():
    from ..colors import ColorMapping
    return ColorMapping


class TestColorMapping:

    def test_ctor(self, ColorMapping):
        colors = ColorMapping()
        xml = tostring(colors.to_tree())
        expected = """
        <clrMapOvr accent1="accent1" accent2="accent2"
           accent3="accent3" accent4="accent4" accent5="accent5"
           accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink"
           hlink="hlink" tx1="dk1" tx2="dk2"
        />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ColorMapping):
        src = """
        <clrMapOvr accent1="accent1" accent2="accent2"
           accent3="accent3" accent4="accent4" accent5="accent5"
           accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink"
           hlink="hlink" tx1="dk1" tx2="dk2"
        />
        """
        node = fromstring(src)
        colors = ColorMapping.from_tree(node)
        assert colors == ColorMapping()
