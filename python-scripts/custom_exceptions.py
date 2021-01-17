"""Custom exceptions thrown by the Python scripts."""

__author__ = "Christopher Menon"
__credits__ = "Christopher Menon"
__license__ = "gpl-3.0"


class ConversionTimeoutException(Exception):
    """Thrown when the PDF to XLSX conversion times out."""

    def __init__(self, converter: str = None, timeout: int = None,
                 message: str = None):
        """Constructs the message using the converter and timeout.

        :param converter: the name of the converter
        :type converter: str, optional
        :param timeout: time waited in seconds
        :type timeout: int, optional
        :param message: a custom message
        :type message: str, optional
        """

        if message is None and isinstance(converter, str) and \
                isinstance(timeout, int):
            self.message = "Waited %d seconds for file conversion from %s." \
                           % (timeout, converter)
        else:
            self.message = message
        super().__init__(self.message)


class ConversionRejectedException(Exception):
    """Thrown when the PDF to XLSX conversion is rejected."""

    def __init__(self, converter: str = None, message: str = None):
        """Constructs the message using the converter.

        :param converter: the name of the converter
        :type converter: str, optional
        :param message: a custom message
        :type message: str, optional
        """

        if message is None and isinstance(converter, str):
            self.message = "The request to convert the PDF to XLSX was " \
                           "rejected by %s." % converter
        else:
            self.message = message
        super().__init__(self.message)


class AppsScriptApiException(Exception):
    """Thrown when the Apps Script API returns an error."""
    pass
