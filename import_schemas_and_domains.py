import os
from openpyxl import load_workbook
import arcpy as ap

# Paths
out_gdb = r'\\spatialfiles.bcgov\work\env\esd\eis\tei\TEI_Working\cmacnaug\TEI_InfoRequests\Amy_Waterhouse_2024_04_19\test5.gdb'
out_folder_path = os.path.dirname(out_gdb)
out_gdb_name = os.path.basename(out_gdb)
schema_excel = r"\\spatialfiles.bcgov\work\env\esd\eis\tei\TEI_Working\cmacnaug\TEI_InfoRequests\Amy_Waterhouse_2024_04_19\schemas.xlsx"
domains_excel = r"\\spatialfiles.bcgov\work\env\esd\eis\tei\TEI_Working\cmacnaug\TEI_InfoRequests\Amy_Waterhouse_2024_04_19\domains.xlsx"

try:
    # Create geodatabase if it doesn't exist
    if not os.path.exists(out_gdb):
        ap.management.CreateFileGDB(out_folder_path, out_gdb_name)
        print(f'\nGeodatabase created: {out_gdb}')

    # Set workspace
    ap.env.workspace = out_gdb
    ap.env.overwriteOutput = True

    # Function to import tables from Excel
    def import_tables_from_excel(excel_file):
        workbook = load_workbook(filename=excel_file)
        sheets = workbook.sheetnames
        print(f'{len(sheets)} found: {sheets}')
        for sheet in sheets:
            out_table = os.path.join(out_gdb, sheet)
            ap.conversion.ExcelToTable(excel_file, out_table, sheet)

    # Import schema tables
    print('\nImporting schema tables...')
    import_tables_from_excel(schema_excel)

    # Create feature classes from tables
    for table in ap.ListTables():
        out_fc = table + "_FC"
        ap.management.CreateFeatureclass(out_gdb, out_fc, "POINT")
        ap.management.JoinField(out_fc, "OBJECTID", table, "OBJECTID")
        ap.management.DeleteRows(out_fc)
        ap.management.Delete(table)
        print(f'{out_fc} feature class created')

    # Import domain tables and assign domains
    print('\nImporting domain tables...')
    import_tables_from_excel(domains_excel)
    for sheet in ap.ListTables():
        ap.management.TableToDomain(sheet, "Code", "Description", out_gdb, sheet)
        ap.management.Delete(sheet)
        print(f'{sheet} domain created')

    # Assign domains to feature classes
    for fc in ap.ListFeatureClasses():
        for field in ap.ListFields(fc):
            for domain in ap.da.ListDomains():
                if field.name == domain.name:
                    ap.management.AssignDomainToField(fc, field.name, domain.name)
                    print(f'{domain.name} domain assigned to field')

    print('\n>>> DONE >>>')

except Exception as e:
    print(f'An error occurred: {e}')
