from __future__ import annotations

import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from xlsheet2csv import cli
from xlsheet2csv.converter import InputLimits

NON_FINITE = ("nan", "NaN", "NAN", "inf", "Inf", "INF", "-inf", "-Infinity", "Infinity", "1e309", "-1e309")


class InputLimitsValidationTests(unittest.TestCase):
    def test_defaults_are_valid(self) -> None:
        InputLimits().validate()

    def test_finite_positive_ratios_are_valid(self) -> None:
        for value in (0.5, 1.0, 250.0, 1e300):
            with self.subTest(value=value):
                InputLimits(max_compression_ratio=value).validate()

    def test_non_finite_ratios_are_rejected(self) -> None:
        for value in (float("nan"), float("inf"), float("-inf")):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, "finite"):
                InputLimits(max_compression_ratio=value).validate()

    def test_zero_and_negative_ratios_are_rejected(self) -> None:
        for value in (0.0, -1.0):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, "greater than zero"):
                InputLimits(max_compression_ratio=value).validate()

    def test_other_limits_are_still_checked(self) -> None:
        with self.assertRaisesRegex(ValueError, "max_sheets"):
            InputLimits(max_sheets=0).validate()


class CliCompressionRatioTests(unittest.TestCase):
    def run_cli(self, *arguments: str) -> tuple[int, str, Path | None, mock.Mock]:
        with tempfile.TemporaryDirectory() as directory:
            output_root = Path(directory) / "out"
            stderr = io.StringIO()
            with mock.patch.object(cli, "convert_path") as convert, contextlib.redirect_stderr(stderr):
                try:
                    code = cli.main([directory, "-o", str(output_root), *arguments])
                except SystemExit as exit_request:
                    code = int(exit_request.code or 0)
            return code, stderr.getvalue(), output_root if output_root.exists() else None, convert

    def test_non_finite_values_are_usage_errors(self) -> None:
        for value in NON_FINITE:
            with self.subTest(value=value):
                code, _, created, convert = self.run_cli(f"--max-compression-ratio={value}")
                self.assertEqual(code, 2)
                self.assertIsNone(created)
                convert.assert_not_called()

    def test_zero_and_negative_values_are_usage_errors(self) -> None:
        for value in ("0", "-1", "-0.5"):
            with self.subTest(value=value):
                code, _, created, convert = self.run_cli(f"--max-compression-ratio={value}")
                self.assertEqual(code, 2)
                self.assertIsNone(created)
                convert.assert_not_called()

    def test_finite_positive_value_reaches_conversion(self) -> None:
        manifest = {"status": "succeeded", "success_count": 0, "failure_count": 0, "failures": []}
        with tempfile.TemporaryDirectory() as directory:
            with mock.patch.object(cli, "convert_path", return_value=(manifest, 0, Path(directory))) as convert:
                with contextlib.redirect_stdout(io.StringIO()):
                    code = cli.main([directory, "--max-compression-ratio=100.5"])
        self.assertEqual(code, 0)
        self.assertEqual(convert.call_args.kwargs["limits"].max_compression_ratio, 100.5)

    def test_default_ratio_is_unchanged(self) -> None:
        args = cli.build_parser().parse_args(["source"])
        self.assertEqual(args.max_compression_ratio, 250.0)


if __name__ == "__main__":
    unittest.main()
