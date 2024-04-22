# -*- coding: utf-8 -*-

import os
from openpyxl import load_workbook
import arcpy as ap


class Toolbox:
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Import Schemas and Domains"
        self.alias = "ImportSchemasAndDomains"

        # List of tool classes associated with this toolbox
        self.tools = [ImportSchemasAndDomains]


class ImportSchemasAndDomains:
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Import Schemas and Domains"
        self.description = """This tool takes input schema and domain Excel files, 
                            producing a feature class or domain table for each sheet within them.
                            The domains are then assigned to the appropriate fields within the feature classes."""

    def getParameterInfo(self):
        """Define the tool parameters."""
        param0 = ap.Parameter(
        displayName="File Geodatabase",
        name="gdb",
        datatype="DEWorkspace",
        parameterType="Required",
        direction="Output")

        param0.filter.list = ['gdb']

        param1 = ap.Parameter(
        displayName="Input Schemas Excel",
        name="in_schemas",
        datatype="DEFile",
        parameterType="Required",
        direction="Input")

        param1.filter.list = ['xlsx']

        param2 = ap.Parameter(
        displayName="Input Domains Excel",
        name="in_domains",
        datatype="DEFile",
        parameterType="Required",
        direction="Input")

        param2.filter.list = ['xlsx']

        parameters = [param0, param1, param2]
        return parameters

    def isLicensed(self):
        """Set whether the tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter. This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        # Paths
        out_gdb = parameters[0].valueAsText
        out_folder_path = os.path.dirname(out_gdb)
        out_gdb_name = os.path.basename(out_gdb)
        schema_excel = parameters[1].valueAsText
        domains_excel = parameters[2].valueAsText

        try:
            # Create geodatabase if it doesn't exist
            if not os.path.exists(out_gdb):
                ap.management.CreateFileGDB(out_folder_path, out_gdb_name)
                messages.addMessage(f'\nGeodatabase created: {out_gdb}')

            # Set workspace
            ap.env.workspace = out_gdb
            ap.env.overwriteOutput = True

            # Function to import tables from Excel
            def import_tables_from_excel(excel_file):
                workbook = load_workbook(filename=excel_file)
                sheets = workbook.sheetnames
                messages.addMessage(f'{len(sheets)} found: {sheets}')
                for sheet in sheets:
                    out_table = os.path.join(out_gdb, sheet)
                    ap.conversion.ExcelToTable(excel_file, out_table, sheet)

            # Import schema tables
            messages.addMessage('\nImporting schema tables...')
            import_tables_from_excel(schema_excel)

            # Create feature classes from tables
            for table in ap.ListTables():
                out_fc = table + "_FC"
                ap.management.CreateFeatureclass(out_gdb, out_fc, "POINT")
                ap.management.JoinField(out_fc, "OBJECTID", table, "OBJECTID")
                ap.management.DeleteRows(out_fc)
                ap.management.Delete(table)
                messages.addMessage(f'{out_fc} feature class created')

            # Import domain tables and assign domains
            messages.addMessage('\nImporting domain tables...')
            import_tables_from_excel(domains_excel)
            for sheet in ap.ListTables():
                ap.management.TableToDomain(sheet, "Code", "Description", out_gdb, sheet)
                ap.management.Delete(sheet)
                messages.addMessage(f'{sheet} domain created')

            # Assign domains to feature classes
            for fc in ap.ListFeatureClasses():
                for field in ap.ListFields(fc):
                    for domain in ap.da.ListDomains():
                        if field.name == domain.name:
                            ap.management.AssignDomainToField(fc, field.name, domain.name)
                            messages.addMessage(f'{domain.name} domain assigned to field')

            messages.addMessage('\n>>> DONE >>>')

        except Exception as e:
            messages.addErrorMessage(f'An error occurred: {e}')

        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return
