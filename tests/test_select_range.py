import unittest
import pandas as pd
import numpy as np
import xlgrab


class TestSelectRange(unittest.TestCase):

    def setUp(self):
        # 5x5 示例表
        self.df = xlgrab.XlDataFrame({
            'A': ['A1','A2','A3','A4','A5'],
            'B': ['B1','B2','B3','B4','B5'],
            'C': ['C1','C2','C3','C4','C5'],
            'D': ['D1','D2','D3','D4','D5'],
            'E': ['E1','E2','E3','E4','E5'],
        })

    def test_cell_to_cell(self):
        out = self.df.select_range(start='A2', end=('cell','C4'))
        self.assertEqual(out.shape, (3, 3))
        self.assertEqual(out.iloc[0,0], 'A2')
        self.assertEqual(out.iloc[-1,-1], 'C4')

    def test_row_and_col_specs(self):
        # 行指定到表尾，列用字母
        out = self.df.select_range(start_row=('row', 2), end_row='end', start_col='B', end_col='D')
        # 行: 2..5 共4行；列: B..D 共3列
        self.assertEqual(out.shape, (4, 3))
        self.assertEqual(out.iloc[0,0], 'B2')
        self.assertEqual(out.iloc[-1,-1], 'D5')

    def test_find_row_and_col(self):
        # 通过列名在列上找行边界；通过行索引在行上找列边界
        # name 行边界用 'A' 列含 'A3' 与 'A4'
        out = self.df.select_range(
            start_row=('find-row', 'A', 'A3', {'mode': 'exact', 'nth': 1}),
            end_row=('find-row', 'A', 'A4', {'mode': 'exact', 'nth': 1}),
            start_col=('find-col', 0, '^B', {'mode': 'regex', 'nth': 1}),
            end_col=('find-col', 0, '^D', {'mode': 'regex', 'nth': 1}),
        )
        self.assertEqual(out.shape, (2, 3))
        self.assertEqual(out.iloc[0,0], 'B3')
        self.assertEqual(out.iloc[-1,-1], 'D4')

    def test_unified_offset_with_clip(self):
        out = self.df.select_range(start='A2', end=('cell','C4'), offset_rows=1, offset_cols=-1, clip=True)
        # 原 A2:C4 偏移到 行3..5，列 A..B → 3行2列
        self.assertEqual(out.shape, (3, 2))
        self.assertEqual(out.iloc[0,0], 'A3')
        self.assertEqual(out.iloc[-1,-1], 'B5')

    def test_flexible_offset_no_clip_error(self):
        with self.assertRaises(ValueError):
            self.df.select_range(
                start='B2', end=('cell','C3'),
                offset_start_row=-10, offset_end_row=0,
                offset_start_col=0, offset_end_col=0,
                clip=False,
            )

    def test_default_bounds_and_order_swap(self):
        # 只给 start，end 用默认末端，且 start/end 交换应自动纠正
        out = self.df.select_range(start='C4', end='A2')
        self.assertEqual(out.iloc[0,0], 'A2')
        self.assertEqual(out.iloc[-1,-1], 'C4')


if __name__ == '__main__':
    unittest.main()


