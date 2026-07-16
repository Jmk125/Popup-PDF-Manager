"""Regression tests for startup-safe helpers."""

import io
import os
import tempfile
import unittest

import pdf_merge_popup as app


class LoggingTests(unittest.TestCase):
    def test_log_works_when_windowed_build_has_no_stdout(self):
        """PyInstaller --noconsole sets sys.stdout to None on Windows."""
        original_path = app.LOG_PATH
        original_stdout = app.sys.stdout
        with tempfile.TemporaryDirectory() as directory:
            app.LOG_PATH = os.path.join(directory, "debug.log")
            app.sys.stdout = None
            try:
                app.log("startup message")
            finally:
                app.sys.stdout = original_stdout
                app.LOG_PATH = original_path

            with open(os.path.join(directory, "debug.log"), encoding="utf-8") as fh:
                self.assertIn("startup message", fh.read())

    def test_log_still_writes_to_an_available_console(self):
        original_path = app.LOG_PATH
        original_stdout = app.sys.stdout
        console = io.StringIO()
        with tempfile.TemporaryDirectory() as directory:
            app.LOG_PATH = os.path.join(directory, "debug.log")
            app.sys.stdout = console
            try:
                app.log("console message")
            finally:
                app.sys.stdout = original_stdout
                app.LOG_PATH = original_path

        self.assertIn("console message", console.getvalue())


if __name__ == "__main__":
    unittest.main()
