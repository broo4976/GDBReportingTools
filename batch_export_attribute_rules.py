"""
Script Name: batch_export_attribute_rules.py
Author: Brooke Reams (breams@esri.com)
Date: May 14, 2026

Description:
    Exports attribute rules, in batch, for all feature classes and tables
    in the input geodatabase or feature dataset.

Inputs:
    - in_ws (str): Path to a geodatabase or feature dataset.

    - out_fldr (str): Path to the folder where the attribute rule csv files
        will be saved.

    - attr_rule_types (list):

Outputs:

Notes:


Versions:
    - ArcGIS Pro 3.6
    - Python 3.11.11

Copyright (c) 2026 Esri. All rights reserved.

Updates:
5/29/20256:     Added new param to allow user to choose specific attribute types to export.

"""

import arcpy
import os
import datetime
import pandas as pd


def log_it(message, level=0):
    print(message)
    if level == 0:
        arcpy.AddMessage(message)
    elif level == 1:
        arcpy.AddWarning(message)
    else:
        arcpy.AddError(message)


def export_attr_rules(ds, attr_rule_types, now):
    # Check if feature class has attribute rules
    if arcpy.Describe(ds).attributeRules:
        if not attr_rule_types:
            log_it(f"Exporting attribute rules for: {ds}")
            out_file = os.path.join(out_fldr, f"{ds}.csv")
        else:
            out_file = os.path.join(out_fldr, f"{ds}_{now}.csv")
        arcpy.management.ExportAttributeRules(ds, out_file)
        if attr_rule_types:
            # Split attribute rules by type
            df = pd.read_csv(out_file)
            for ar_type in attr_rule_types:
                df_filter = df[df["TYPE"] == ar_type.upper()]
                if not df_filter.empty:
                    filter_out_file = os.path.join(out_fldr, f"{ds}_{ar_type}.csv")
                    log_it(f"Exporting {ar_type.upper()} Rules for: {ds}")
                    df_filter.to_csv(filter_out_file, index=False)
            # Delete original out file
            os.remove(out_file)
            xml_file = out_file.replace(".csv", ".csv.xml")
            os.remove(xml_file)


# Inputs
in_ws = arcpy.GetParameterAsText(0)
out_fldr = arcpy.GetParameterAsText(1)
attr_rule_types = arcpy.GetParameter(2)

# Set the workspace environment
arcpy.env.workspace = in_ws

# Get current date/time to use for temp attribute rule files
now = datetime.datetime.now().strftime("%Y%m%d%H%M%S")

# Check if input workspace is geodatabase or feature dataset
data_type = arcpy.Describe(in_ws).dataType
log_it(f"Input workspace type: {data_type.replace("Workspace", "Geodatabase")}")
# If datatype is a file gdb, get all feature datasets
if data_type == "Workspace":
    fds_list = arcpy.ListDatasets(feature_type="Feature")
    for fds in fds_list:
        for fc in arcpy.ListFeatureClasses(feature_dataset=fds):
            export_attr_rules(fc, attr_rule_types, now)

# Get stand-alone feature classes and tables
ds_list = arcpy.ListFeatureClasses() + arcpy.ListTables()
for ds in ds_list:
    export_attr_rules(ds, attr_rule_types, now)
