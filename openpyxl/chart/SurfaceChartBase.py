
#Autogenerated schema
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,)


class BandFmt(Serialisable):

    idx = Typed(expected_type=UnsignedInt, )
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)

    __elements__ = ('idx', 'spPr')

    def __init__(self,
                 idx=None,
                 spPr=None,
                ):
        self.idx = idx
        self.spPr = spPr


class BandFmts(Serialisable):

    bandFmt = Typed(expected_type=BandFmt, allow_none=True)

    __elements__ = ('bandFmt',)

    def __init__(self,
                 bandFmt=None,
                ):
        self.bandFmt = bandFmt


class SurfaceSer(Serialisable):

    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('cat', 'val', 'extLst')

    def __init__(self,
                 cat=None,
                 val=None,
                 extLst=None,
                ):
        self.cat = cat
        self.val = val
        self.extLst = extLst


class SurfaceChartShared(Serialisable):

    wireframe = Bool(nested=True, allow_none=True)
    ser = Typed(expected_type=SurfaceSer, allow_none=True)
    bandFmts = Typed(expected_type=BandFmts, allow_none=True)

    __elements__ = ('wireframe', 'ser', 'bandFmts')

    def __init__(self,
                 wireframe=None,
                 ser=None,
                 bandFmts=None,
                ):
        self.wireframe = wireframe
        self.ser = ser
        self.bandFmts = bandFmts

