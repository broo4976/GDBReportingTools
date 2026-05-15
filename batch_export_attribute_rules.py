"""
Script Name: batch_export_attribute_rules.py
Author: Brooke Reams (breams@esri.com)
Date: May 14, 2026

Description:
    Exports attribute rules, in batch, for all feature classes and tables
    in the input geodatabase or feature dataset.

Inputs:
    - in_ws (str): Path to a geodatabase or feature dataset.

Outputs:
    - out_fldr (str): Path to the folder where the attribute rule csv files
    will be saved.

Notes:


Versions:
    - ArcGIS Pro 3.6
    - Python 3.11.11

Copyright (c) 2026 Esri. All rights reserved.

Updates:

"""

import arcpy
import os


def log_it(message, level=0):
    print(message)
    if level == 0:
        arcpy.AddMessage(message)
    elif level == 1:
        arcpy.AddWarning(message)
    else:
        arcpy.AddError(message)


# Inputs
in_ws = arcpy.GetParameterAsText(0)
out_fldr = arcpy.GetParameterAsText(1)

# Set the workspace environment
arcpy.env.workspace = in_ws

# Check if input workspace is geodatabase or feature dataset
data_type = arcpy.Describe(in_ws).dataType
log_it(f"Input workspace type: {data_type.replace("Workspace", "Geodatabase")}")
# If datatype is a file gdb, get all feature datasets
if data_type == "Workspace":
    fds_list = arcpy.ListDatasets(feature_type="Feature")
    for fds in fds_list:
        for fc in arcpy.ListFeatureClasses(feature_dataset=fds):
            # Check if feature class has attribute rules
            if arcpy.Describe(fc).attributeRules:
                log_it(f"Exporting attribute rules for: {fc}")
                out_file = os.path.join(out_fldr, f"{fc}.csv")
                arcpy.management.ExportAttributeRules(fc, out_file)

# Get stand-alone feature classes and tables
ds_list = arcpy.ListFeatureClasses() + arcpy.ListTables()
for ds in ds_list:
    if arcpy.Describe(ds).attributeRules:
        log_it(f"Exporting attribute rules for: {ds}")
        out_file = os.path.join(out_fldr, f"{ds}.csv")
        arcpy.management.ExportAttributeRules(ds, out_file)
