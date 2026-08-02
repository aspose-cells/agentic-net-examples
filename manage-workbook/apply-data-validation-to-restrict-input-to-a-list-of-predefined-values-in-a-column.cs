using System;
using Aspose.Cells;

namespace AsposeCellsDataValidationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range for the validation (e.g., column A rows 1 to 100)
            // CellArea uses zero‑based indexes: row 0 = first row, column 0 = column A
            CellArea validationArea = CellArea.CreateCellArea(0, 0, 99, 0);

            // Add a new validation to the worksheet for the defined area
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Set the validation type to List and provide the allowed values
            validation.Type = ValidationType.List;
            validation.Formula1 = "Option1,Option2,Option3";

            // Enable the in‑cell drop‑down so users can pick from the list
            validation.InCellDropDown = true;

            // Optional: show an input message when the cell is selected
            validation.ShowInput = true;
            validation.InputTitle = "Select an option";
            validation.InputMessage = "Please choose one of the predefined options.";

            // Optional: show an error message if an invalid value is entered
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid entry";
            validation.ErrorMessage = "The value must be one of the predefined options.";
            validation.AlertStyle = ValidationAlertType.Stop;

            // Save the workbook to a file
            workbook.Save("ColumnListValidation.xlsx");
        }
    }
}