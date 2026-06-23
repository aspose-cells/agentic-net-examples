using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the validation area for column K (zero‑based column index 10)
            // Here we apply the validation to rows 0 through 1048575 (the maximum Excel rows)
            CellArea validationArea = CellArea.CreateCellArea(0, 10, 1048575, 10);

            // Add a new validation to the worksheet for the defined area
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation to allow only whole numbers between 0 and 500
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "0";
            validation.Formula2 = "500";

            // Optional: set user‑visible messages
            validation.ShowInput = true;
            validation.InputTitle = "Enter a number";
            validation.InputMessage = "Please enter a whole number between 0 and 500.";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "The value must be a whole number between 0 and 500.";
            validation.AlertStyle = ValidationAlertType.Stop;

            // Save the workbook with the validation applied
            workbook.Save("ColumnKValidation.xlsx");
        }
    }
}