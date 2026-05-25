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
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the validation range for column T (zero‑based column index 19)
            // Here we apply it to the whole column from row 0 to the maximum row count.
            int maxRow = worksheet.Cells.MaxDataRow; // Use existing max row (0 if empty)
            // If the sheet is empty, set a reasonable range, e.g., first 1000 rows
            if (maxRow < 0) maxRow = 999;
            CellArea timeRange = CellArea.CreateCellArea(0, 19, maxRow, 19);

            // Add a new validation to the worksheet
            ValidationCollection validations = worksheet.Validations;
            int validationIndex = validations.Add(timeRange);
            Validation validation = validations[validationIndex];

            // Configure the validation to allow only time values between 09:00 and 17:00
            validation.Type = ValidationType.Time;                     // Time validation
            validation.Operator = OperatorType.Between;                // Between operator
            validation.Formula1 = "09:00";                             // Lower bound
            validation.Formula2 = "17:00";                             // Upper bound

            // Optional: user-friendly messages
            validation.InputTitle = "Allowed Time";
            validation.InputMessage = "Enter a time between 09:00 and 17:00.";
            validation.ErrorTitle = "Invalid Time";
            validation.ErrorMessage = "The time must be between 09:00 and 17:00.";
            validation.ShowInput = true;
            validation.ShowError = true;
            validation.AlertStyle = ValidationAlertType.Stop;          // Prevent invalid entry

            // Save the workbook
            workbook.Save("TimeValidationColumnT.xlsx");
        }
    }
}