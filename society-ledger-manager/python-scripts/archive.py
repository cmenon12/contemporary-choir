"""Archive of code that's no longer used.

This contains code that is no longer actively used, but may be of use in
the future. Note that its dependencies are not necessarily listed in
requirements.txt.

Currently this only contains the PDFtoXLSXConverter class, which can
convert the PDF ledger to an XLSX file using pdftoexcel.com or
pdftoexcelconverter.net. It also contains ConversionTimeoutError and
ConversionRejectedError which this class uses.
"""

import json
import logging
import random

import requests

from ledger_fetcher import Ledger, CustomEncoder

__author__ = "Christopher Menon"
__credits__ = "Christopher Menon"
__license__ = "gpl-3.0"

# The number of PDF to XLSX converters in the PDFtoXLSXConverter class
NUMBER_OF_CONVERTERS = 2

# How long to wait for conversion (in seconds)
CONVERSION_TIMEOUT = 120


class PDFToXLSXConverter:
    """Represents a PDF to XLSX converter."""

    def __init__(self, ledger: Ledger, converter_number: int = 0):
        """Creates the converter.
        :param ledger: the ledger to convert
        :type ledger: Ledger
        :param converter_number: the chosen converter
        :type converter_number: int, optional
        """

        # Define the user agent, which doesn't change
        user_agent = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/87.0.4280.141 Safari/537.36 Edg/87.0.664.75")

        # Define the fixed headers
        headers = {
            "Connection": "keep-alive",
            "X-Requested-With": "XMLHttpRequest",
            "DNT": "1",
            "User-Agent": user_agent,
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Dest": "empty",
            "Accept-Language": "en-GB,en;q=0.9",
        }

        # If no valid number was specified then choose randomly
        if converter_number not in range(1, NUMBER_OF_CONVERTERS + 1):
            converter_number = random.randint(1, NUMBER_OF_CONVERTERS)

        # Use pdftoexcel.com
        if converter_number == 1:
            self.name = "pdftoexcel.com"
            self.request_url = "https://www.pdftoexcel.com/upload.instant.php"
            headers["Accept"] = "*/*"
            headers["Origin"] = "https://www.pdftoexcel.com"
            headers["Referer"] = "https://www.pdftoexcel.com/"
            self.files = {"Filedata": open(ledger.get_pdf_filepath(), "rb")}
            self.status_url = "https://www.pdftoexcel.com/status"
            self.download_url = "https://www.pdftoexcel.com"

        # Use pdftoexcelconverter.net
        else:
            self.name = "pdftoexcelconverter.net"
            self.request_url = "https://www.pdftoexcelconverter.net/upload.instant.php"
            headers["Accept"] = "application/json"
            headers["Origin"] = "https://www.pdftoexcelconverter.net"
            headers["Referer"] = "https://www.pdftoexcelconverter.net/"
            self.files = {"file[0]": open(ledger.get_pdf_filepath(), "rb")}
            self.status_url = "https://www.pdftoexcelconverter.net/getIsConverted.php"
            self.download_url = "https://www.pdftoexcelconverter.net"

        # Create the session that'll be used to make all the requests
        self.session = requests.Session()
        self.session.headers.update(headers)

        self.log()

    def upload_pdf(self, raise_for_status: bool = True) -> str:
        """Make the conversion request and upload the PDF.
        :param raise_for_status: whether to raise_for_status()
        :type raise_for_status: bool, optional
        :return: the conversion job ID
        :rtype: str
        :raises HTTPError: if a bad HTTP status code is returned
        :raises ConversionRejectedError: if server rejects the PDF
        """

        response = self.session.post(url=self.request_url, files=self.files)
        if raise_for_status:
            response.raise_for_status()
        if "jobId" not in response.json().keys():
            raise ConversionRejectedError(self.get_name())
        return response.json()["jobId"]

    def check_conversion_status(self, job_id: str,
                                raise_for_status: bool = True) \
            -> str:
        """Check the status of the request and get the download URL.
        :param job_id: the ID of the conversion job
        :type job_id: str
        :param raise_for_status: whether to raise_for_status()
        :type raise_for_status: bool, optional
        :return: the download URL (which might be empty)
        :rtype: str
        :raises HTTPError: if a bad HTTP status code is returned
        :raises JSONDecodeError: if the response can't be decoded
        """

        response = self.session.get(url=self.status_url,
                                    params=(("jobId", job_id), ("rand", "16")))
        if raise_for_status:
            response.raise_for_status()
        return response.json()["download_url"]

    def download_xlsx(self, job_id: str, download_url: str,
                      raise_for_status: bool = True) -> bytes:
        """Download the XLSX file.
        :param job_id: the ID of the conversion job
        :type job_id: str
        :param download_url: the download URL
        :type download_url: str
        :param raise_for_status: whether to raise_for_status()
        :type raise_for_status: bool, optional
        :return: the XLSX file
        :rtype: bytes
        :raises HTTPError: if a bad HTTP status code is returned
        """

        response = self.session.get(url=self.download_url + download_url,
                                    params={"id": job_id})
        if raise_for_status:
            response.raise_for_status()
        return response.content

    def get_name(self) -> str:
        """Return the name of the converter."""
        return self.name

    def log(self) -> None:
        """Logs the object to the log."""

        LOGGER.info(json.dumps(self.__dict__, cls=CustomEncoder))


class ConversionTimeoutError(Exception):
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


class ConversionRejectedError(Exception):
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


if __name__ == "__main__":

    # Prepare the log
    logging.basicConfig(filename="ledger_fetcher.log",
                        filemode="a",
                        format="%(asctime)s | %(levelname)s : %(message)s",
                        level=logging.DEBUG)
    LOGGER = logging.getLogger(__name__)

else:
    LOGGER = logging.getLogger(__name__)
