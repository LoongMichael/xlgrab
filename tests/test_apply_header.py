import unittest
import pandas as pd

from xlgrab.core import XlDataFrame


class TestApplyHeader(unittest.TestCase):

    def setUp(self):
        # 基础样例：前两行为表头，后两行为数据
        self.raw = pd.DataFrame({
            'A': ['H1', 'H2', 'r1', 'r2'],
            'B': ['C1', 'C2', 1, 2],
            'C': ['X1', 'X2', 3, 4],
        })
        self.df = XlDataFrame(self.raw)

    def test_header_true_single_row(self):
        out = self.df.apply_header(True)
        self.assertEqual(out.shape, (3, 3))
        self.assertEqual(out.columns.tolist(), ['H1', 'C1', 'X1'])

    def test_header_two_rows_default_join_underscore(self):
        # 与 read_csv 语义：多行用 list[int]
        out = self.df.apply_header([0,1])
        self.assertEqual(out.shape, (2, 3))
        self.assertEqual(out.columns.tolist(), ['H1_H2', 'C1_C2', 'X1_X2'])

    def test_header_two_rows_multiindex(self):
        out = self.df.apply_header([0,1], header_join=None)
        self.assertEqual(out.shape, (2, 3))
        self.assertTrue(isinstance(out.columns, pd.MultiIndex))
        self.assertEqual(list(out.columns.map(lambda t: tuple(t))), [('H1', 'H2'), ('C1', 'C2'), ('X1', 'X2')])

    def test_header_list_with_duplicates(self):
        out = self.df.apply_header(['a', 'a', 'a'])
        self.assertEqual(out.columns.tolist(), ['a', 'a_1', 'a_2'])

    def test_header_series(self):
        s = pd.Series(['col', 'col', 'col'])
        out = self.df.apply_header(s)
        self.assertEqual(out.columns.tolist(), ['col', 'col_1', 'col_2'])

    def test_header_dataframe_join(self):
        header_df = pd.DataFrame({
            'A': ['Top', 'Sub'],
            'B': ['Alpha', 'Beta'],
            'C': ['I', 'II']
        })
        out = self.df.apply_header(header_df)  # 默认 '_'
        self.assertEqual(out.columns.tolist(), ['Top_Sub', 'Alpha_Beta', 'I_II'])

    def test_safe_name_and_dedup_integration(self):
        raw2 = pd.DataFrame({
            'A': ['Name (CN)', '子 类', 'v1'],
            'B': ['Amt-USD', '金 额', 'v2'],
            'C': ['Date/Time', '日 期', 'v3'],
            'D': ['Name (CN)', '子 类', 'v4'],
        })
        df2 = XlDataFrame(raw2)
        out = df2.apply_header([0,1])  # 使用默认下划线合并 + 规范化 + 去重
        # 规范化后 Name (CN) -> Name_CN，子 类 -> 子_类 等，并处理重复列名
        self.assertEqual(out.columns.tolist(), ['Name_CN_子_类', 'Amt_USD_金_额', 'Date_Time_日_期', 'Name_CN_子_类_1'])

    def test_dedup_many(self):
        out = self.df.apply_header(['x', 'x', 'x', ])
        self.assertEqual(out.columns.tolist(), ['x', 'x_1', 'x_2'])


if __name__ == '__main__':
    unittest.main()


