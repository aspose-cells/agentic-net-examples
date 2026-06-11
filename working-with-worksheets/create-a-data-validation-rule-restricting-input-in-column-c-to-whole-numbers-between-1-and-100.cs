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

            // Define the validation area for column C (zero‑based column index 2)
            // Here we apply the rule to rows 0 through 1000; adjust as needed.
            CellArea area = CellArea.CreateCellArea(0, 2, 1000, 2);

            // Add a new validation to the worksheet for the defined area
            int validationIndex = worksheet.Validations.Add(area);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation: whole numbers between 1 and 100
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";

            // Optional: user-friendly messages
            validation.InputTitle = "Enter Whole Number";
            validation.InputMessage = "Please enter a whole number between 1 and 100.";
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "The value must be a whole number between 1 and 100.";
            validation.ShowInput = true;
            validation.ShowError = true;

            // Save the workbook
            workbook.Save("ColumnC_WholeNumber_1_to_100.xlsx");
        }
    }
}