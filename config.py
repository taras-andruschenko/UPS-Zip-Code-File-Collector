import os

DOWNLOAD_URL = "https://www.ups.com/media/us/currentrates/zone-csv/{}.xls"

FILES_DIRECTORY = "zip_files"
ROOT_DIRECTORY = os.getcwd()
PATH_TO_SAVE_FILES = os.path.join(ROOT_DIRECTORY, FILES_DIRECTORY)

ZONE_RANGES_FILE_NAME = "Carriers zone ranges.xlsx"
COLUMN_FROM = "zip from"
COLUMN_TO = "zip to"
SHEET = "UPS zip ranges"
