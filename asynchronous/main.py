import asyncio
import aiohttp
import os
import time
import pandas as pd

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
        str(worksheet[column_from].loc[row] + 1).zfill(5):
            str(worksheet[column_to].loc[row]).zfill(5)
        for row in range(len(worksheet))
    }


def get_current_codes(file: str) -> tuple[str, str]:
    """
    This function will extract the last zip code of the current
    """
    # Read current zones source file
    worksheet = pd.read_excel(
        os.path.join(ROOT_DIRECTORY, FILES_DIRECTORY, file),
        sheet_name=0,
        header=None
    )

    # Get a cell that contains information about
    # the initial and final zip codes.
    phrase = str(worksheet.loc[4][0])

    # Reformat given text according to zip codes representation
    last_code = phrase[
                phrase.find("to ") + 3: phrase.rfind(".")
                ].replace("-", "")
    first_code = phrase[
                 phrase.find("to ") - 7: phrase.find("to ") - 1
                 ].replace("-", "")

    return first_code, last_code


async def get_file(zip_code: str, session) -> str | None:
    """
    This function will download single domestic zones source file
    with custom zip code. It will return file's name.
    """

    # Zip code must be truncated to the first 3 characters.
    # This is required to generate the correct download URL.
    prefix = zip_code[:3]

    # Create and write a new file with context manager
    # and send a request inside it.
    with open(
            os.path.join(PATH_TO_SAVE_FILES, "{}.xlsx".format(prefix)),
            "wb"
    ) as file:
        response = await session.get(
                url=DOWNLOAD_URL.format(prefix),
                ssl=False
        )
        data = await response.read()
        file.write(data)

        return file.name


async def get_all_files(zip_ranges: dict) -> None:
    """
    This function will collect all domestic zones source files.
    """
    connector = aiohttp.TCPConnector()

    # Get an async session.
    async with aiohttp.ClientSession(connector=connector) as session:

        # Get all available files in the zip_ranges dict in async mode
        tasks = []
        for zip_code in zip_ranges:
            tasks.append(asyncio.ensure_future(get_file(zip_code, session)))

        files = await asyncio.gather(*tasks)

        # Check the file. If the last zip code doesn't match
        # the value from the zip_ranges dictionary, assign new
        # initial zip code and download the file with the missing range.
        #
        # Also, please, note that this block of code may be and even
        # must be rewritten asynchronously!
        #
        # For example, without the block code below,
        # the script execution time takes about 10 seconds!!!
        for file in files:
            current_first_code, current_last_code = get_current_codes(file)
            expected_last_code = zip_ranges[current_first_code]

            while int(current_last_code) < int(expected_last_code):

                current_first_code = str(int(current_last_code) + 1)

                new_file = await get_file(current_first_code, session)
                _, current_last_code = get_current_codes(new_file)


def main() -> None:

    # Note: ZONE_RANGES_FILE_NAME, COLUMN_FROM, COLUMN_TO and
    # SHEET may be reassigned in config.py according to your initial data!
    zip_ranges_dict = get_zip_ranges_dict(
        ZONE_RANGES_FILE_NAME,
        COLUMN_FROM,
        COLUMN_TO,
        SHEET,
    )
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(get_all_files(zip_ranges_dict))


# __main__ stamp will be deleted!
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
