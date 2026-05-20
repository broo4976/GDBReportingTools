"""
Script Name: batch_import_attribute_rules.py
Author: Brooke Reams (breams@esri.com)
Date: May 14, 2026

Description:
    Imports attribute rules, in batch, to the input geodatabase.

Inputs:
    - in_fldr (str): Path to folder containing attribute rule csv files.

Outputs:
    - in_gdb (str): Path to the geodatabase where the attribute rules will
    be imported.

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
in_gdb = arcpy.GetParameterAsText(0)
in_fldr = arcpy.GetParameterAsText(1)

# Set the workspace envionment
arcpy.env.workspace = in_gdb

# Get dictionary of feature classes/tables
log_it("Reading input workspace")
ds_dict = {}  # {fc: fds, fc: fds, fc: "", tbl: ""}
fds_list = arcpy.ListDatasets(feature_type="Feature")
fds_list.insert(0, "")
for fds in fds_list:
    for fc in arcpy.ListFeatureClasses(feature_dataset=fds):
        ds_dict[fc] = fds

# Get a list of csv files from input folder
log_it("Getting attribute rule csv files from input folder")
csv_list = [f for f in os.listdir(in_fldr) if f.lower().endswith(".csv")]

# Loop through csv files and find corresponding fc/table
log_it("Looping through attribute rule csv files")
for csv_file in csv_list:
    # Get name of file
    name = csv_file.split(".")[0]
    if name in ds_dict.keys():
        fds = ds_dict[name]
        ds = os.path.join(in_gdb, fds, name)
        # Import attribute rules
        log_it(f"Importing {csv_file} to {ds}")
        arcpy.management.ImportAttributeRules(ds, os.path.join(in_fldr, csv_file))
