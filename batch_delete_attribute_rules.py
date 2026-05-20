"""
Script Name: batch_delete_attribute_rules.py
Author: Brooke Reams (breams@esri.com)
Date: May 15, 2026

Description:
    Deletes attribute rules, in batch, for all feature classes and tables
    in the input geodatabase or feature dataset.

Inputs:
    - in_ws (str): Path to a geodatabase or feature dataset containing attribute
        rules to delete.

Outputs:
    - log_flie (str): Optional file containing log messages.

Notes:


Versions:
    - ArcGIS Pro 3.6
    - Python 3.11.11

Copyright (c) 2026 Esri. All rights reserved.

Updates:

"""

import arcpy
import logging


def log_it(message):
    print(message)
    arcpy.AddMessage(message)

    if out_log:
        logging.info(message)


# Inputs
in_ws = arcpy.GetParameterAsText(0)
out_log = arcpy.GetParameterAsText(1)

# Set the workspace environment
arcpy.env.workspace = in_ws

# Create log file
if out_log:
    logging.basicConfig(
        filename=out_log,
        filemode="w",
        format="%(asctime)s - %(message)s",
        datefmt="%d-%b-%y %H:%M:%S",
        level=logging.INFO,
        force=True,
    )

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
                # Get attribute rule
                attr_rules = arcpy.Describe(fc).attributeRules
                # Loop through attribute rules and delete
                attr_names = [ar.name for ar in attr_rules]
                log_it(f"Deleting {len(attr_names)} attribute rules from {fc}")
                arcpy.management.DeleteAttributeRule(fc, attr_names)

# Get stand-alone feature classes and tables
ar_found = False
ds_list = arcpy.ListFeatureClasses() + arcpy.ListTables()
for ds in ds_list:
    # Get attribute rule
    attr_rules = arcpy.Describe(ds).attributeRules
    if attr_rules:
        ar_found = True
        # Get list of attribute rules and delete
        attr_names = [ar.name for ar in attr_rules]
        log_it(f"Deleting {len(attr_names)} attribute rules from {ds}")
        arcpy.management.DeleteAttributeRule(ds, attr_names)

if not ar_found:
    log_it(
        "Attribute rules were not found on any datasets within the provided workspace"
    )
