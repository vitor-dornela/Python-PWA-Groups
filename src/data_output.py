import pandas as pd
import logging
from .utils import extraction_complete_message

def save_to_excel(groups: list, users: list, categories: list, output_file: str):
    df_groups = pd.DataFrame(groups)
    df_users = pd.DataFrame(users)
    df_categories = pd.DataFrame(categories).sort_values(by="Category Name")

    with pd.ExcelWriter(output_file) as writer:
        df_users.to_excel(writer, sheet_name="Users", index=False)
        df_groups.to_excel(writer, sheet_name="Groups", index=False)
        df_categories.to_excel(writer, sheet_name="Categories", index=False)
    print(extraction_complete_message(output_file))
