"""
multi_n_export.py
-----------------
Kept for backward compatibility. The Multi N Excel exporter now lives in
gap_excel_export.py together with the main all-types exporter.
"""

from gap_excel_export import export_multi_n_excel  # noqa: F401

__all__ = ["export_multi_n_excel"]
