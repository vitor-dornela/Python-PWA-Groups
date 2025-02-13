import sys
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# User Constants
FILE_NAME = "pwa_data"
OUTPUT_DIRECTORY = "~/Downloads"

# Program Constants
LOGIN_URL = "https://login.microsoftonline.com/"
GROUP_CONTAINER_ID = "GridDataRow"
USER_CONTAINER_ID = "ctl00_ctl00_PlaceHolderMain_PWA_PlaceHolderMain_idFormSectionUsers_ctl02_idSwpUsers_BetaList_Container"
CATEGORY_CONTAINER_ID = "ctl00_ctl00_PlaceHolderMain_PWA_PlaceHolderMain_idFormSectionCategories_ctl02_idSwpCategories_BetaList_Container"
CHROME_TIMEOUT = 10

sys.stdout.reconfigure(encoding='utf-8')
