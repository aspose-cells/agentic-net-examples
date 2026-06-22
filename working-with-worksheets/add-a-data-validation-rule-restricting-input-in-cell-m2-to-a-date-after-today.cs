using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the cell area for M2 (row index 1, column index 12)
        CellArea area = CellArea.CreateCellArea(1, 12, 1, 12);

        // Add a validation to the worksheet for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Configure the validation: Date type, must be greater than TODAY()
        validation.Type = ValidationType.Date;
        validation.Operator = OperatorType.GreaterThan;
        validation.Formula1 = "TODAY()";

        // Optional: display input message when the cell is selected
        validation.ShowInput = true;
        validation.InputTitle = "Date Validation";
        validation.InputMessage = "Enter a date later than today.";

        // Optional: display error message for invalid input
        validation.ShowError = true;
        validation.ErrorTitle = "Invalid Date";
        validation.ErrorMessage = "The date must be after today.";

        // Save the workbook
        workbook.Save("DateValidationAfterToday.xlsx");
    }
}