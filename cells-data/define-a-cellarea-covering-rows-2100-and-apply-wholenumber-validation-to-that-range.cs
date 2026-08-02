using System;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define a cell area that covers rows 2‑100 in column A (zero‑based indices)
            // Row 2 -> index 1, Row 100 -> index 99, Column A -> index 0
            CellArea validationArea = CellArea.CreateCellArea(1, 0, 99, 0);

            // Add a whole‑number validation to the defined area
            int validationIndex = sheet.Validations.Add(validationArea);
            Validation validation = sheet.Validations[validationIndex];

            // Configure the validation as a whole‑number between 1 and 1000
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";      // Minimum allowed value
            validation.Formula2 = "1000";   // Maximum allowed value
            validation.ShowInput = true;
            validation.ShowError = true;
            validation.InCellDropDown = false;

            // Optionally, put some sample data in the range
            for (int row = 1; row <= 99; row++)
            {
                cells[row, 0].PutValue(row); // Example values
            }

            // Save the workbook
            workbook.Save("WholeNumberValidation.xlsx", SaveFormat.Xlsx);
        }
    }
}