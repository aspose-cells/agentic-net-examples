using System;
using Aspose.Cells;

class DataValidationExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Custom formula validation: cell A1 must be greater than the value in B1
        // -------------------------------------------------
        // Define the validation area (A1)
        CellArea customArea = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 0, EndColumn = 0 };
        int customIndex = sheet.Validations.Add(customArea);
        Validation customValidation = sheet.Validations[customIndex];

        // Set validation type to Custom and operator to GreaterThan
        customValidation.Type = ValidationType.Custom;
        customValidation.Operator = OperatorType.GreaterThan;

        // Set the formula referencing B1 (A1 notation, not R1C1, not locale-specific)
        customValidation.SetFormula1("=B1", false, false);

        // Configure error display
        customValidation.ShowError = true;
        customValidation.ErrorTitle = "Invalid Input";
        customValidation.ErrorMessage = "Value must be greater than the value in B1.";

        // Provide a sample value in B1 for the formula to reference
        sheet.Cells["B1"].PutValue(10);

        // -------------------------------------------------
        // List validation: cells C1:C5 allow only specific options
        // -------------------------------------------------
        // Define the validation area (C1:C5)
        CellArea listArea = new CellArea { StartRow = 0, StartColumn = 2, EndRow = 4, EndColumn = 2 };
        int listIndex = sheet.Validations.Add(listArea);
        Validation listValidation = sheet.Validations[listIndex];

        // Set validation type to List and enable the drop‑down
        listValidation.Type = ValidationType.List;
        listValidation.InCellDropDown = true;

        // Provide a comma‑separated list of allowed values
        listValidation.Formula1 = "OptionA,OptionB,OptionC";

        // Configure error display
        listValidation.ShowError = true;
        listValidation.ErrorTitle = "Selection Error";
        listValidation.ErrorMessage = "Please select a value from the list.";

        // Save the workbook to an XLSX file
        workbook.Save("DataValidationExample.xlsx");
    }
}