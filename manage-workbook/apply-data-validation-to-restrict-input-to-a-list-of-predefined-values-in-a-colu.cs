using System;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range for the validation (e.g., column B rows 1 to 100)
            // Rows and columns are zero‑based indexes
            int startRow = 0;      // Row 1
            int endRow = 99;       // Row 100
            int columnIndex = 1;   // Column B
            CellArea validationArea = CellArea.CreateCellArea(startRow, columnIndex, endRow, columnIndex);

            // Add a new validation to the worksheet for the defined area
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation to be a list with predefined values
            validation.Type = ValidationType.List;
            validation.Formula1 = "Option1,Option2,Option3,Option4";
            validation.InCellDropDown = true;   // Show drop‑down arrow in the cell
            validation.ShowInput = true;        // Optional: show input message
            validation.InputTitle = "Select Value";
            validation.InputMessage = "Please choose one of the predefined options.";
            validation.ShowError = true;        // Optional: show error on invalid entry
            validation.ErrorTitle = "Invalid Selection";
            validation.ErrorMessage = "The value you entered is not in the allowed list.";

            // Save the workbook to a file
            workbook.Save("ColumnValidationDemo.xlsx");
        }
    }
}