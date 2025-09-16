import pandas as pd
import logging
from .utils import extraction_complete_message

def _create_table_and_format(worksheet, df, table_name, table_style):
    """Helper function to create table and format columns."""
    if df.empty:
        return
    
    # Create Excel Table (Insert > Table equivalent)
    from openpyxl.worksheet.table import Table, TableStyleInfo
    
    # Define table range
    end_column = chr(64 + len(df.columns))  # Convert to letter (A, B, C...)
    table_range = f"A1:{end_column}{len(df) + 1}"
    
    # Create table with style
    table = Table(displayName=table_name, ref=table_range)
    style = TableStyleInfo(
        name=table_style, 
        showFirstColumn=False,
        showLastColumn=False, 
        showRowStripes=True, 
        showColumnStripes=False
    )
    table.tableStyleInfo = style
    
    # Add table to worksheet
    worksheet.add_table(table)
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

def save_to_excel(groups: list, users: list, categories: list, output_file: str):
    logging.info("Salvando dados em Excel: %d grupos, %d usuários, %d categorias", len(groups), len(users), len(categories))
    df_groups = pd.DataFrame(groups)
    df_users = pd.DataFrame(users)
    df_categories = pd.DataFrame(categories).sort_values(by="Category Name") if categories else pd.DataFrame()

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Save and format Users sheet
        if not df_users.empty:
            df_users.to_excel(writer, sheet_name="Users", index=False)
            _create_table_and_format(writer.sheets["Users"], df_users, "UsersTable", "TableStyleMedium4")

        # Save and format Groups sheet
        if not df_groups.empty:
            df_groups.to_excel(writer, sheet_name="Groups", index=False)
            _create_table_and_format(writer.sheets["Groups"], df_groups, "GroupsTable", "TableStyleMedium6")

        # Save and format Categories sheet
        if not df_categories.empty:
            df_categories.to_excel(writer, sheet_name="Categories", index=False)
            _create_table_and_format(writer.sheets["Categories"], df_categories, "CategoriesTable", "TableStyleMedium7")

    print(extraction_complete_message(output_file))
