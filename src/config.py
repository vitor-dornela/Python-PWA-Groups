import sys
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# User Constants
FILE_NAME = "pwa_data"
OUTPUT_DIRECTORY = "~/Downloads"

# Authentication and Navigation Constants
LOGIN_URL = "https://login.microsoftonline.com/"
MANAGE_USERS_PATH = "_layouts/15/pwa/Admin/ManageUsers.aspx"
BROWSER_TIMEOUT = 10

# PWA Group Extraction Constants
GROUP_CONTAINER_ID = "GridDataRow"
USER_CONTAINER_ID = "ctl00_ctl00_PlaceHolderMain_PWA_PlaceHolderMain_idFormSectionUsers_ctl02_idSwpUsers_BetaList_Container"
CATEGORY_CONTAINER_ID = "ctl00_ctl00_PlaceHolderMain_PWA_PlaceHolderMain_idFormSectionCategories_ctl02_idSwpCategories_BetaList_Container"

# PWA User Grid Extraction Constants
USERS_GRID_ID = "ctl00_ctl00_PlaceHolderMain_PWA_PlaceHolderMain_idGrdUsers"

# PWA Export Column Mapping (exact order from table structure)
PWA_COLUMNS = {
    0: 'Checkbox',  # Skip checkbox column
    1: 'User Name',
    2: 'Email Address', 
    3: 'User Logon Account',
    4: 'State',
    5: 'RBS'
}

sys.stdout.reconfigure(encoding='utf-8')
