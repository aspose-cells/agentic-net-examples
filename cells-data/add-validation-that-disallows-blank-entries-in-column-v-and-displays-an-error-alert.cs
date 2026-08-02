using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the validation area for column V (index 21) from row 0 to row 1000
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 1000,
            StartColumn = 21, // Column V
            EndColumn = 21
        };

        // Add a validation to the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Configure validation to disallow blank entries
        validation.Type = ValidationType.AnyValue;      // Accept any value
        validation.IgnoreBlank = false;                // Do NOT ignore blanks (i.e., blanks are invalid)

        // Set error alert properties
        validation.AlertStyle = ValidationAlertType.Stop; // Show a Stop alert
        validation.ErrorTitle = "Invalid Input";
        validation.ErrorMessage = "Blank entries are not allowed in column V.";
        validation.ShowError = true;                     // Display the error message when validation fails

        // Save the workbook
        workbook.Save("ColumnV_NoBlank_Validation.xlsx");
    }
}