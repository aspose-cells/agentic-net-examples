using System;
using Aspose.Cells;

namespace EmailUniqueValidation
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range for the email column (AK = column index 36)
            // Assuming data starts at row 2 (index 1) and goes down to row 1000
            int startRow = 1;          // Row 2 in Excel (0‑based index)
            int endRow = 1000;         // Adjust as needed
            int emailColumn = 36;      // Column AK (0‑based)

            // Create a validation that checks each email is unique within the range
            CellArea emailRange = CellArea.CreateCellArea(startRow, emailColumn, endRow, emailColumn);
            int validationIndex = worksheet.Validations.Add(emailRange);
            Validation validation = worksheet.Validations[validationIndex];

            // Set validation type to Custom and provide a formula that returns TRUE for unique values
            validation.Type = ValidationType.Custom;
            // Formula uses absolute range for COUNTIF and relative reference for the current cell
            validation.Formula1 = $"=COUNTIF($AK${startRow + 1}:$AK${endRow + 1},AK{startRow + 1})=1";

            // Configure the error alert that will be shown for duplicate entries
            validation.ShowError = true;
            validation.AlertStyle = ValidationAlertType.Stop;
            validation.ErrorTitle = "Invalid Email";
            validation.ErrorMessage = "Duplicate email address not allowed.";

            // Save the workbook with the validation applied
            workbook.Save("output.xlsx");
        }
    }
}