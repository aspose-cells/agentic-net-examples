using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the validation area for column E (zero‑based column index 4)
            // Here we apply the validation to rows 0‑1000; adjust as needed.
            CellArea validationArea = CellArea.CreateCellArea(0, 4, 1000, 4);

            // Add a new validation to the worksheet
            ValidationCollection validations = worksheet.Validations;
            int validationIndex = validations.Add(validationArea);
            Validation validation = validations[validationIndex];

            // Configure the validation to allow only dates between Jan 1 and Dec 31
            validation.Type = ValidationType.Date;                     // Date validation
            validation.Operator = OperatorType.Between;                // Between two dates
            validation.Formula1 = "1/1/2023";                           // Lower bound (Jan 1)
            validation.Formula2 = "12/31/2023";                         // Upper bound (Dec 31)

            // Optional user‑friendly messages
            validation.InputTitle = "Date Required";
            validation.InputMessage = "Enter a date between Jan 1 and Dec 31.";
            validation.ErrorTitle = "Invalid Date";
            validation.ErrorMessage = "The date must be within the year 2023.";
            validation.ShowInput = true;
            validation.ShowError = true;

            // Save the workbook
            workbook.Save("ColumnEDateValidation.xlsx");
        }
    }
}