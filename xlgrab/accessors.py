"""
Pandas DataFrame/Series accessors for xlgrab.

Usage:
    import xlgrab  # ensures accessor registration via __init__
    import pandas as pd
    df = pd.DataFrame(...)
    df.xl.excel_range("A1:B2")

Optionally, users can enable direct methods (df.excel_range) via
xlgrab.enable_direct_methods() if they desire.
"""

from pandas.api.extensions import register_dataframe_accessor, register_series_accessor
import pandas as pd

# Import the implementation from our enhanced classes to reuse logic
from .core import XlDataFrame, XlSeries


@register_dataframe_accessor("xl")
class XlDataFrameAccessor:
    def __init__(self, pandas_obj: pd.DataFrame):
        self._obj = pandas_obj

    # Delegate to XlDataFrame methods by binding self._obj as "self"
    def excel_range(self, *args, **kwargs):
        return XlDataFrame.excel_range(self._obj, *args, **kwargs)

    def offset_range(self, *args, **kwargs):
        return XlDataFrame.offset_range(self._obj, *args, **kwargs)

    def select_range(self, *args, **kwargs):
        return XlDataFrame.select_range(self._obj, *args, **kwargs)

    def find_idx(self, *args, **kwargs):
        return XlDataFrame.find_idx(self._obj, *args, **kwargs)

    def apply_header(self, *args, **kwargs):
        return XlDataFrame.apply_header(self._obj, *args, **kwargs)


@register_series_accessor("xl")
class XlSeriesAccessor:
    def __init__(self, pandas_obj: pd.Series):
        self._obj = pandas_obj

    def find_idx(self, *args, **kwargs):
        return XlSeries.find_idx(self._obj, *args, **kwargs)


def enable_direct_methods() -> None:
    """Optionally attach direct-call methods to pandas DataFrame/Series.

    After calling this, users can use df.excel_range(...) directly.
    This does not replace classes; it only sets attributes on pd.DataFrame/Series.
    """
    # DataFrame direct bindings
    pd.DataFrame.excel_range = XlDataFrame.excel_range
    pd.DataFrame.offset_range = XlDataFrame.offset_range
    pd.DataFrame.select_range = XlDataFrame.select_range
    pd.DataFrame.find_idx = XlDataFrame.find_idx
    pd.DataFrame.apply_header = XlDataFrame.apply_header

    # Series direct bindings
    pd.Series.find_idx = XlSeries.find_idx


