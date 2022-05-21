"""Custom exceptions thrown by the Python scripts."""

__author__ = "Christopher Menon"
__credits__ = "Christopher Menon"
__license__ = "gpl-3.0"


class AppsScriptApiError(Exception):
    """Thrown when the Apps Script API returns an error."""
    pass


class XLSXDoesNotExistError(Exception):
    """Thrown when the XLSX file doesn't exist."""
    pass


class URLDoesNotExistError(Exception):
    """Thrown when a URL doesn't exist."""
    pass


class PDFIsNotSavedError(Exception):
    """Thrown when the PDF file isn't saved."""
    pass


class XLSXIsNotSavedError(Exception):
    """Thrown when the XLSX file isn't saved."""
    pass
