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

def save_to_excel(groups: list, users_groups: list, categories: list, all_users_detailed: list, output_file: str):
    logging.info("Salvando dados em Excel: %d grupos, %d usuários (grupos), %d categorias, %d usuários (detalhados)", 
                len(groups), len(users_groups), len(categories), len(all_users_detailed))
    
    df_groups = pd.DataFrame(groups)
    df_users_groups = pd.DataFrame(users_groups)
    df_categories = pd.DataFrame(categories).sort_values(by="Category Name") if categories else pd.DataFrame()
    df_all_users = pd.DataFrame(all_users_detailed) if all_users_detailed else pd.DataFrame()

    # Create CountUsersGroups aggregation based on users_groups
    df_count_users_groups = pd.DataFrame()
    if not df_users_groups.empty:
        # Count users per group
        user_count_by_group = df_users_groups.groupby('Group Name').size().reset_index(name='UsersCount')
        user_count_by_group = user_count_by_group.sort_values(by='Group Name')
        df_count_users_groups = user_count_by_group

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Save and format new Users sheet (detailed info from Excel export)
        if not df_all_users.empty:
            df_all_users.to_excel(writer, sheet_name="Users", index=False)
            _create_table_and_format(writer.sheets["Users"], df_all_users, "UsersTable", "TableStyleMedium2")

        # Save and format UsersGroups sheet (user-group relationships)
        if not df_users_groups.empty:
            df_users_groups.to_excel(writer, sheet_name="UsersGroups", index=False)
            _create_table_and_format(writer.sheets["UsersGroups"], df_users_groups, "UsersGroupsTable", "TableStyleMedium4")

        # Save and format CountUsersGroups sheet (aggregation)
        if not df_count_users_groups.empty:
            df_count_users_groups.to_excel(writer, sheet_name="CountUsersGroups", index=False)
            _create_table_and_format(writer.sheets["CountUsersGroups"], df_count_users_groups, "CountUsersGroupsTable", "TableStyleMedium5")

        # Save and format Groups sheet
        if not df_groups.empty:
            df_groups.to_excel(writer, sheet_name="Groups", index=False)
            _create_table_and_format(writer.sheets["Groups"], df_groups, "GroupsTable", "TableStyleMedium6")

        # Save and format Categories sheet
        if not df_categories.empty:
            df_categories.to_excel(writer, sheet_name="Categories", index=False)
            _create_table_and_format(writer.sheets["Categories"], df_categories, "CategoriesTable", "TableStyleMedium7")

    print(extraction_complete_message(output_file))
