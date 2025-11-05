"""
Brooke Reams - breams@esri.com
Oct. 16, 2025

Description:
Loops through all feature classes and tables at the root
and feature dataset level and reports basic properties
on each dataset.

ArcGIS Pro 3.5.2
Python 3.11.11

Updates:
11/4/2025:      Fix for tables; changed MakeFeatureLayer to MakeTableView.

"""

import arcpy
import os
import openpyxl
from openpyxl.utils import get_column_letter

# Overwrite existing output
arcpy.env.overwriteOutput = 1


def log_it(message):
    print(message)
    arcpy.AddMessage(message)


def autofit_column_widths(ws):
    for col in ws.columns:
        max_length = 0
        column = get_column_letter(
            col[0].column
        )  # Get column letter from the first cell in the column
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except (
                TypeError
            ):  # Handle cases where cell.value might be None or not easily convertible to string
                pass
        adjusted_width = max_length + 2  # Add some padding
        ws.column_dimensions[column].width = adjusted_width


# Tool inputs
in_ws = arcpy.GetParameterAsText(0)
out_xls = arcpy.GetParameterAsText(1)

# Create new workbook and define header
wb = openpyxl.Workbook()
ws = wb.active
ws["A1"] = "Feature Dataset"
ws["B1"] = "Feature Class/Table"
ws["C1"] = "Shape Type"
ws["D1"] = "Record Count"
ws["E1"] = "Subtype Name"
ws["F1"] = "Subtype Count"


# Set workspace environment
arcpy.env.workspace = in_ws

# Initialize list to store data properties
records = []

# Loop through all feature classes/tables in gdb and in all feature datasets
fds_list = arcpy.ListDatasets(feature_type="Feature")
fds_list.sort()
fds_list.append("")
for fds in fds_list:
    # Get feature classes/tables
    if fds == "":
        log_it(f"Processing stand-alone datasets")
        fds = "<standalone>"
        ds_list = arcpy.ListFeatureClasses() + arcpy.ListTables()
    else:
        log_it(f"Processing feature dataset: {fds}")
        ds_list = arcpy.ListFeatureClasses(feature_dataset=fds)

    for ds in ds_list:
        log_it(f"Processing dataset: {ds}")
        # Get record count
        record_count = int(arcpy.management.GetCount(ds).getOutput(0))
        # Describe dataset to get properties
        desc = arcpy.Describe(ds)
        # Get shape type
        try:
            shape_type = desc.shapeType
        except:
            shape_type = ""
            pass
        # Get subtype info
        subtype_fld = desc.subtypeFieldName
        subtype_dict = arcpy.da.ListSubtypes(ds)
        subtype_list = []
        for i, subtype_prop in subtype_dict.items():
            if i > 0:
                subtype_name = subtype_prop["Name"]
                # Get count of records for subtype
                where_clause = f"{subtype_fld} = {i}"
                arcpy.management.MakeTableView(ds, "ds_lyr", where_clause)
                subtype_count = int(arcpy.management.GetCount("ds_lyr").getOutput(0))
                subtype_list.append((subtype_name, subtype_count))
        # Add details to data list
        val_tuple = (fds, ds, shape_type, record_count, subtype_list)
        records.append(val_tuple)
    records.append("")

# Write results to excel
if records:
    row = 2
    for val in records:
        if val == "":
            row += 1
        else:
            for i in range(0, len(val) - 1):
                ws.cell(row=row, column=i + 1, value=val[i])
            row += 1
            if val[-1]:
                for subtype in val[-1]:
                    ws.cell(row=row, column=5, value=subtype[0])
                    ws.cell(row=row, column=6, value=subtype[1])
                    row += 1

    # Update formatting for record count columns
    for cell in ws["D"]:
        cell.number_format = "#,##0"

    for cell in ws["F"]:
        cell.number_format = "#,##0"

    # Bold and freeze first row
    bold_font = openpyxl.styles.Font(bold=True)
    for cell in ws[1]:
        cell.font = bold_font
    ws.freeze_panes = "A2"

    # Apply autofit to all columns
    autofit_column_widths(ws)

    # Save excel
    wb.save(out_xls)

    # Start file
    os.startfile(out_xls)
