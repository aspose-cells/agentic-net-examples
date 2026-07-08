using System;
using Aspose.Cells;

namespace AsposeCellsDateValidation
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the validation range for column N (index 13)
            // Here we apply the validation to rows 1 through 1000 in column N.
            CellArea dateRange = CellArea.CreateCellArea(0, 13, 999, 13); // A1-style: N1:N1000

            // Add a new validation to the worksheet for the defined range
            int validationIndex = sheet.Validations.Add(dateRange);
            Validation validation = sheet.Validations[validationIndex];

            // Configure the validation to allow dates between 01/01/2020 and 12/31/2025
            validation.Type = ValidationType.Date;                     // Date validation
            validation.Operator = OperatorType.Between;                // Between operator
            validation.Formula1 = new DateTime(2020, 1, 1).ToOADate().ToString(); // Lower bound
            validation.Formula2 = new DateTime(2025, 12, 31).ToOADate().ToString(); // Upper bound

            // Optional: user-friendly messages
            validation.InputTitle = "Enter a date";
            validation.InputMessage = "Date must be between 01/01/2020 and 12/31/2025.";
            validation.ErrorTitle = "Invalid Date";
            validation.ErrorMessage = "The entered date is outside the allowed range.";
            validation.ShowInput = true;
            validation.ShowError = true;

            // Save the workbook
            workbook.Save("DateValidationColumnN.xlsx");
        }
    }
}