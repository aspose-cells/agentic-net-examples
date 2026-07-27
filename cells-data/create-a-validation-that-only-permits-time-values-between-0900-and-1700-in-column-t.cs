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

            // Define the validation range for column T (column index 19)
            // Apply to all rows from 1 to the maximum row count in Excel
            CellArea timeRange = CellArea.CreateCellArea("T1", "T1048576");

            // Add a new validation to the worksheet for the defined range
            int validationIndex = worksheet.Validations.Add(timeRange);
            Validation timeValidation = worksheet.Validations[validationIndex];

            // Set validation type to Time and require values between 09:00 and 17:00
            timeValidation.Type = ValidationType.Time;
            timeValidation.Operator = OperatorType.Between;
            // Excel stores time as a fraction of a day; using TIME function ensures correct serial values
            timeValidation.Formula1 = "=TIME(9,0,0)";   // 09:00
            timeValidation.Formula2 = "=TIME(17,0,0)"; // 17:00

            // Optional: user-friendly messages
            timeValidation.InputTitle = "Allowed Time";
            timeValidation.InputMessage = "Enter a time between 09:00 and 17:00.";
            timeValidation.ErrorTitle = "Invalid Time";
            timeValidation.ErrorMessage = "The time must be between 09:00 and 17:00.";
            timeValidation.AlertStyle = ValidationAlertType.Stop;
            timeValidation.ShowInput = true;
            timeValidation.ShowError = true;

            // Save the workbook
            workbook.Save("TimeValidationColumnT.xlsx");
        }
    }
}