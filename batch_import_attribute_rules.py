"""
Script Name: batch_import_attribute_rules.py
Author: Brooke Reams (breams@esri.com)
Date: May 14, 2026

Description:
    Imports attribute rules, in batch, to the input geodatabase or feature
        dataset.

Inputs:
    - in_gdb (str): Path to the geodatabase or feature dataset where the
        attribute rules will be imported.

    - in_fldr (str): Path to folder containing attribute rule csv files.

    - attr_rule_types (list): List of attribute ruletypes to import.  If an
        attribute rule type is selected, the tool will look specifially for any
        csv files that match the name of the datasets with the attribute type
        appended to the name.  For example, if Validation is selected, the tool
        will look specifically for csv files that match the following pattern:
        <dataset name>_Validation.csv.  If no attribute rule types are selected,
        the tool will look for csv files that match the dataset name only:
        <dataset name>.csv.

Outputs:


Notes:


Versions:
    - ArcGIS Pro 3.6
    - Python 3.11.11

Copyright (c) 2026 Esri. All rights reserved.

Updates:
5/29/20256:     Added new param to allow user to choose specific attribute types to import.

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
attr_rule_types = arcpy.GetParameter(2)

# Set the workspace envionment
arcpy.env.workspace = in_gdb

# Get dictionary of feature classes/tables
log_it("Reading input workspace")
ds_dict = {}  # {fc: fds, fc: fds, fc: "", tbl: ""}
case_dict = {}  # {fc: FC, fc: Fc}
fds_list = arcpy.ListDatasets(feature_type="Feature")
fds_list.insert(0, "")
for fds in fds_list:
    for fc in arcpy.ListFeatureClasses(feature_dataset=fds):
        ds_dict[fc.lower()] = fds
        case_dict[fc.lower()] = fc

# Get a list of csv files from input folder
log_it("Getting attribute rule csv files from input folder")
csv_all_list = [f for f in os.listdir(in_fldr) if f.lower().endswith(".csv")]
csv_filter_list = []
if attr_rule_types:
    for ar_type in attr_rule_types:
        for csv_file in csv_all_list:
            if csv_file.lower().endswith(f"{ar_type.lower()}.csv"):
                csv_filter_list.append(csv_file)
else:
    for csv_file in csv_all_list:
        if csv_file.lower().split("_")[-1] not in [s.lower() for s in attr_rule_types]:
            csv_filter_list.append(csv_file)


# Loop through csv files and find corresponding fc/table
log_it("Looping through attribute rule csv files")
for csv_file in csv_filter_list:
    # Get name of file
    file_name_no_ext = csv_file.lower().split(".csv")[0]
    # Get feature class name
    if file_name_no_ext.lower().split("_")[-1] in [s.lower() for s in attr_rule_types]:
        name = file_name_no_ext.replace(
            f"_{file_name_no_ext.lower().split("_")[-1]}", ""
        )
    else:
        name = file_name_no_ext.lower()

    if name in ds_dict.keys():
        fds = ds_dict[name]
        ds = os.path.join(in_gdb, fds, case_dict[name])
        # Import attribute rules
        log_it(f"Importing {csv_file} to {ds}")
        arcpy.management.ImportAttributeRules(ds, os.path.join(in_fldr, csv_file))
        if arcpy.GetMessages(1):
            log_it(arcpy.GetMessages(1), 1)
