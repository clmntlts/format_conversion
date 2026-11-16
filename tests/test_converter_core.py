"""Unit tests for converter_core.  Run with:  python -m unittest"""

import os
import sys
import tempfile
import unittest

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import converter_core as cc  # noqa: E402


def long_rows(rows):
    cols = ["submission_id", "question_code", "question_type", "matrix_row",
            "answer_1", "answer_2", "answer_3"]
    return pd.DataFrame(rows, columns=cols)


class ConverterCoreTests(unittest.TestCase):
    def test_pivot_simple_and_matrix(self):
        df = long_rows([
            ("s1", "Q1", "coded_answer", "row_0", "AB", None, None),
            ("s1", "M", "matrix", "a.\tfirst item", "Jamais", None, None),
            ("s1", "M", "matrix", "b) second item", "Souvent", None, None),
            ("s2", "Q1", "coded_answer", "row_0", "CD", None, None),
            ("s2", "M", "matrix", "a.\tfirst item", "Toujours", None, None),
        ])
        wide, cb = cc.to_wide(df)
        self.assertEqual(list(wide.columns), ["submission_id", "Q1", "M_a", "M_b"])
        self.assertEqual(wide.set_index("submission_id").loc["s2", "M_a"], "Toujours")
        self.assertTrue(pd.isna(wide.set_index("submission_id").loc["s2", "M_b"]))
        self.assertEqual(list(cb["variable"]), ["Q1", "M_a", "M_b"])
        self.assertEqual(cb.loc[0, "matrix_row"], "")

    def test_multiple_answers_are_joined(self):
        df = long_rows([("s1", "Q", "qcm", "row_0", "Oui", "Non", None)])
        wide, _ = cc.to_wide(df)
        self.assertEqual(wide.loc[0, "Q"], "Oui | Non")

    def test_row0_and_blank_matrix_row_are_the_same_variable(self):
        df = long_rows([
            ("s1", "Q", "qcm", "row_0", "x", None, None),
            ("s2", "Q", "qcm", None, "y", None, None),
        ])
        wide, cb = cc.to_wide(df)
        self.assertEqual(list(cb["variable"]), ["Q"])
        self.assertEqual(list(wide["Q"]), ["x", "y"])

    def test_integer_answers_survive_excel_round_trip(self):
        df = long_rows([
            ("s1", "Q", "coded_answer", "row_0", 3, None, None),
            ("s2", "Q", "coded_answer", "row_0", None, None, None),
            ("s3", "Q", "coded_answer", "row_0", 2.5, None, None),
        ])
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "in.xlsx")
            df.to_excel(path, index=False)
            wide, _ = cc.to_wide(cc.load_long(path))
        self.assertEqual(list(wide["Q"]), ["3", "", "2.5"])

    def test_label_without_prefix_gets_unique_slug(self):
        df = long_rows([
            ("s1", "M", "matrix", "Item without prefix", "1", None, None),
            ("s1", "M", "matrix", "Item without prefix, again", "2", None, None),
        ])
        _, cb = cc.to_wide(df)
        self.assertEqual(list(cb["variable"]), ["M_Item_without", "M_Item_without_2"])

    def test_data_sheet_detection_and_export(self):
        df = long_rows([("s1", "Q", "qcm", "row_0", "x", None, None)])
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "in.xlsx")
            with pd.ExcelWriter(src) as xw:
                pd.DataFrame({"note": ["summary"]}).to_excel(xw, sheet_name="Summary", index=False)
                df.to_excel(xw, sheet_name="data", index=False)
            sheet, _ = cc.find_data_sheet(src)
            self.assertEqual(sheet, "data")

            out = os.path.join(tmp, "out.xlsx")
            cc.export_wide(*cc.to_wide(cc.load_long(src)), out)
            with pd.ExcelFile(out) as xl:
                self.assertEqual(xl.sheet_names, ["wide", "codebook"])

            bad = os.path.join(tmp, "bad.xlsx")
            pd.DataFrame({"a": [1]}).to_excel(bad, index=False)
            with self.assertRaises(ValueError):
                cc.load_long(bad)


if __name__ == "__main__":
    unittest.main()
