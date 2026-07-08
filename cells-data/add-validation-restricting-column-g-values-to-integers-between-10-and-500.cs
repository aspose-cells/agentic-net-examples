using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the validation area for column G (zero‑based index 6) covering all rows
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 6,
            EndRow = 1048575,   // maximum row index in Excel
            EndColumn = 6
        };

        // Add a validation to the worksheet for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Configure the validation: whole numbers between 10 and 500
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "10";
        validation.Formula2 = "500";

        // Optional user‑friendly messages
        validation.InputTitle = "Enter Integer";
        validation.InputMessage = "Please enter an integer between 10 and 500.";
        validation.ErrorTitle = "Invalid Input";
        validation.ErrorMessage = "Value must be an integer between 10 and 500.";
        validation.ShowInput = true;
        validation.ShowError = true;
        validation.AlertStyle = ValidationAlertType.Stop;

        // Save the workbook
        workbook.Save("ColumnGValidation.xlsx");
    }
}