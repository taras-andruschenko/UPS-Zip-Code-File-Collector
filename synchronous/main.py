import os
import time
import requests
import pandas as pd

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from config import (
    DOWNLOAD_URL,
    FILES_DIRECTORY,
    ZONE_RANGES_FILE_NAME,
    COLUMN_FROM,
    COLUMN_TO,
    SHEET,
    ROOT_DIRECTORY,
    PATH_TO_SAVE_FILES,
)


def get_session():
    """
    This function will create a session with which you can make
    a large number http requests.
    """

    session = requests.Session()
    retry = Retry(connect=3, backoff_factor=0.5)
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('https://', adapter)

    return session


def get_zip_ranges_dict(
        ranges_file_name: str,
        column_from: str,
        column_to: str,
        sheet: str | int
) -> dict[str]:
    """
    This function will return a dictionary with key as a first zip number
    and value as a last zip number. It will return pairs with five-digit
    strings.
    """

    # Read the source file with the ranges
    worksheet = pd.read_excel(
        os.path.join(ROOT_DIRECTORY, ranges_file_name),
        sheet_name=sheet
    )

    return {
        str(worksheet[column_from].loc[row]).zfill(5):
            str(worksheet[column_to].loc[row]).zfill(5)
        for row in range(len(worksheet))
    }


def get_current_last_code(file: str) -> str:
    """
    This function will extract the last zip code of the current
    """
    worksheet = pd.read_excel(
        os.path.join(ROOT_DIRECTORY, FILES_DIRECTORY, file),
        sheet_name=0,
        header=None
    )
    phrase = str(worksheet.loc[4][0])

    return phrase[phrase.find("to ") + 3: phrase.rfind(".")].replace("-", "")


def get_file(zip_code: str, session) -> str | None:
    """
    This function will download single domestic zones source file
    with custom zip code. It will return file's name.
    """

    # Zip code must be truncated to the first 3 characters.
    # This is required to generate the correct download URL.
    prefix = zip_code[:3]

    # Send a request and handle possible connection errors.
    try:
        response = session.get(
            url=DOWNLOAD_URL.format(prefix)
        )
    except ConnectionError as e:
        print(e, "connections problem")

    # Create and write a new file with data in .xlsx format
    with open(os.path.join(PATH_TO_SAVE_FILES, "{}.xlsx".format(prefix)),
              "wb") as file:
        file.write(response.content)

        return file.name


def get_all_files(zip_ranges: dict) -> str:
    """
    This function will collect all domestic zones source files.
    """

    session = get_session()

    # Get all available files in the zip_ranges dict
    for first_code, last_code in zip_ranges.items():
        file = get_file(first_code, session)
        print(first_code)
        current_last_code = get_current_last_code(file)

        # Check the file. If the last zip code doesn't match
        # the value from the zip_ranges dictionary, assign new
        # initial zip code and download the file with the missing range.
        while int(current_last_code) < int(last_code):

            current_start_code = str(int(current_last_code) + 1)

            new_file = get_file(current_start_code, session)
            current_last_code = get_current_last_code(new_file)

    return "All files have been downloaded successfully"


def main() -> None:

    # Note: ZONE_RANGES_FILE_NAME, COLUMN_FROM, COLUMN_TO and
    # SHEET may be reassigned in config.py according to your initial data!
    zip_ranges_dict = get_zip_ranges_dict(
        ZONE_RANGES_FILE_NAME,
        COLUMN_FROM,
        COLUMN_TO,
        SHEET,
    )
    get_all_files(zip_ranges_dict)


# __main__ stamp will be deleted!
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
