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

    - attr_rule_types (list): Types of attribute rules to delete.  Only types that
        are selected will be deleted.  If no types are selected, the tool will work
        as if all types are checked, and delete all attribute rules for each feature
        class/table within the input workspace.

Outputs:
    - out_log (str): Optional file containing log messages.

Notes:


Versions:
    - ArcGIS Pro 3.6
    - Python 3.11.11

Copyright (c) 2026 Esri. All rights reserved.

Updates:
5/29/20256:     Added new param to allow user to choose specific attribute types to delete.

"""

import arcpy
import logging


def log_it(message):
    print(message)
    arcpy.AddMessage(message)

    if out_log:
        logging.info(message)


def delete_attribute_rules(ds, system_ar_names_list):
    ar_found = False
    # Get attribute rule
    attr_rules = arcpy.Describe(ds).attributeRules
    if attr_rules:
        # Get list of attribute rules and delete
        attr_names = [ar.name for ar in attr_rules if ar.type in system_ar_names_list]
        if attr_names:
            log_it(f"Deleting {len(attr_names)} attribute rule(s) from {ds}")
            arcpy.management.DeleteAttributeRule(ds, attr_names)
            ar_found = True

    return ar_found


# Inputs
in_ws = arcpy.GetParameterAsText(0)
out_log = arcpy.GetParameterAsText(1)
attr_rule_types = arcpy.GetParameter(2)

# Set the workspace environment
arcpy.env.workspace = in_ws

# If no rules were selected in tool ui, add all to list
if not attr_rule_types:
    attr_rule_types = []
    attr_rule_types.append("Constraint")
    attr_rule_types.append("Calculation")
    attr_rule_types.append("Validation")

# Get system names for selected attribute rule types
system_ar_names_dict = {
    "Constraint": "esriARTConstraint",
    "Calculation": "esriARTCalculation",
    "Validation": "esriARTValidation",
}
system_ar_names_list = []
for t in attr_rule_types:
    system_ar_names_list.append(system_ar_names_dict[t])

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
ar_found = []
if data_type == "Workspace":
    fds_list = arcpy.ListDatasets(feature_type="Feature")
    for fds in fds_list:
        for fc in arcpy.ListFeatureClasses(feature_dataset=fds):
            ar_found = delete_attribute_rules(fc, system_ar_names_list)

# Get stand-alone feature classes and tables
ds_list = arcpy.ListFeatureClasses() + arcpy.ListTables()
for ds in ds_list:
    ar_found.append(delete_attribute_rules(ds, system_ar_names_list))

if True not in ar_found:
    log_it(
        f"Attribute rules of type {", ".join(attr_rule_types)} were not found on any datasets within the provided workspace"
    )
