"""Custom exceptions thrown by the Python scripts."""

__author__ = "Christopher Menon"
__credits__ = "Christopher Menon"
__license__ = "gpl-3.0"


class ConversionTimeoutException(Exception):
    """Thrown when the PDF to XLSX conversion times out."""
    pass


class AppsScriptApiException(Exception):
    """Thrown when the Apps Script API returns an error."""
    pass
