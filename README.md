# VBScript Empty Parameter Handling Bug

This repository demonstrates an uncommon bug in VBScript related to handling Empty parameters in functions.  The `IsEmpty()` function is used to check for Empty values and assign defaults (0 in this case), however, this can lead to unexpected results depending on the order of operations and implicit type conversion.