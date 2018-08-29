# generateTestCases

This generates a predefined excel sheet with field level test cases written, by reading a excel sheet which has the following inputs.

Inputs taken for each field:
Field names,
Field type,
Enabled/Disabled condition of the field,

In case enabled,
Maximum Length of the field,
Mandatory condition(Format contains : Y/N),

If field has specific set of values and/or a default value,
Field value column,
Field default value column.

By using this information, the corresponding row numbers and file name should be updated in src/fieldLevelTestcaseWriting/getData.java
UI_SPEC_SHEET_NAME,UI_SPEC_SHEET_NAME, FieldNameColumn, FieldTypeColumn, FieldMaxLengthColumn, FieldMandatoryColumn, EnableDisableColumn, DefaultValueColumn, FieldValueColumn

For the excel file: FieldStartRow, FieldEndRow must be specified.

resources/steps.java should have the steps that are common which are used to guide the user to the fields location.

Note : This just copy pastes the values and can be efficient only if a lot of fields are on the page, and have default functionality.
