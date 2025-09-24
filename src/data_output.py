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
    
    # Auto-fit columns exactly like Excel's double-click behavior
    from openpyxl.utils import get_column_letter
    
    for idx, column in enumerate(worksheet.columns, 1):
        column_letter = get_column_letter(idx)
        
        # Find the maximum content width in this column
        max_width = 0
        for cell in column:
            if cell.value is not None:
                # Convert to string and measure
                cell_text = str(cell.value)
                cell_width = len(cell_text)
                max_width = max(max_width, cell_width)
        
        # Apply Excel-like auto-fit: content width + minimal padding
        if max_width > 0:
            # Excel's auto-fit typically adds ~1.2 character width for padding
            auto_fit_width = max_width + 1.2
            # Set reasonable min/max bounds
            final_width = max(8.43, min(auto_fit_width, 120))  # Excel's default min is ~8.43
        else:
            final_width = 8.43  # Excel's default column width
            
        worksheet.column_dimensions[column_letter].width = final_width

def save_to_excel(groups: list, users_groups: list, categories: list, all_users_detailed: list, output_file: str):
    # Calculate distinct counts for the 3 main data types
    distinct_groups = len(groups)
    distinct_categories = len({cat.get('Category UID') for cat in categories if cat.get('Category UID')})
    distinct_users = len({user.get('User UID') for user in all_users_detailed if user.get('User UID')})
    
    logging.info("Salvando dados em Excel: %d grupos, %d categorias, %d usuários", distinct_groups, distinct_categories, distinct_users)
    
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
        # Always save Users sheet (detailed info from Excel export)
        df_all_users.to_excel(writer, sheet_name="Users", index=False)
        if not df_all_users.empty:
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
