using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the validation range: entire column B (rows 1‑100)
        int startRow = 0;          // Row 1 (zero‑based)
        int endRow = 99;           // Row 100
        int columnIndex = 1;       // Column B (0 = A)

        CellArea area = CellArea.CreateCellArea(startRow, columnIndex, endRow, columnIndex);

        // Add a validation to the worksheet for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Set validation as a list with predefined values
        validation.Type = ValidationType.List;
        validation.Formula1 = "Apple,Banana,Cherry";
        validation.InCellDropDown = true;          // Show drop‑down list in cells
        validation.ShowInput = true;               // Show input message when cell is selected
        validation.InputTitle = "Select Fruit";
        validation.InputMessage = "Choose a fruit from the list.";
        validation.ShowError = true;               // Show error dialog on invalid entry
        validation.ErrorTitle = "Invalid Selection";
        validation.ErrorMessage = "Please select a value from the provided list.";

        // Save the workbook to a file
        workbook.Save("ColumnListValidation.xlsx");
    }
}