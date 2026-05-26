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
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 1. Define a named range that will be used for validation
            // ------------------------------------------------------------
            // Fill some sample data in A1:A5
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue($"Item{i + 1}");
            }

            // Add a named range called "AllowedValues" that refers to A1:A5
            int nameIndex = workbook.Worksheets.Names.Add("AllowedValues");
            Name allowedRange = workbook.Worksheets.Names[nameIndex];
            allowedRange.RefersTo = "=Sheet1!$A$1:$A$5";

            // ------------------------------------------------------------
            // 2. Create a data validation that restricts entry to the named range
            // ------------------------------------------------------------
            // Define the area where the validation will be applied (e.g., B1:B5)
            CellArea validationArea = CellArea.CreateCellArea(0, 1, 4, 1); // rows 0-4, column 1 (B)

            // Add a new validation to the worksheet's validation collection
            int validationIndex = sheet.Validations.Add(validationArea);
            Validation validation = sheet.Validations[validationIndex];

            // Set validation type to List and point Formula1 to the named range
            validation.Type = ValidationType.List;
            // Formula1 must start with '=' and reference the named range
            validation.Formula1 = "=AllowedValues";

            // Optional: show dropdown arrow in the cell
            validation.InCellDropDown = true;

            // Optional: set input and error messages
            validation.ShowInput = true;
            validation.InputTitle = "Select Value";
            validation.InputMessage = "Choose a value from the predefined list.";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Entry";
            validation.ErrorMessage = "Please select a value from the allowed list.";

            // ------------------------------------------------------------
            // 3. Save the workbook
            // ------------------------------------------------------------
            workbook.Save("ValidationWithNamedRange.xlsx");
        }
    }
}