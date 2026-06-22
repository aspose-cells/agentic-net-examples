using System;
using Aspose.Cells;

namespace AsposeCellsValidationAndFreeze
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Add data validation to an entire column (e.g., column B)
            // -------------------------------------------------
            // Define the area: column index 1 (B), rows 0 to 1000
            CellArea validationArea = new CellArea
            {
                StartRow = 0,
                EndRow = 1000,
                StartColumn = 1,
                EndColumn = 1
            };

            // Add the validation to the worksheet's validation collection
            int validationIndex = sheet.Validations.Add(validationArea);
            Validation validation = sheet.Validations[validationIndex];

            // Configure the validation as a drop‑down list
            validation.Type = ValidationType.List;
            validation.InCellDropDown = true;
            validation.Formula1 = "Option1,Option2,Option3";

            // Optional: show input message when the cell is selected
            validation.ShowInput = true;
            validation.InputTitle = "Select an option";
            validation.InputMessage = "Please choose one of the listed values.";

            // -------------------------------------------------
            // 2. Freeze the column so the validation stays visible while scrolling
            // -------------------------------------------------
            // Freeze the first column (A) by setting the freeze point at column B (index 1)
            // Row index = 0 (no rows frozen), column index = 1, freezedRows = 0, freezedColumns = 1
            sheet.FreezePanes(0, 1, 0, 1);

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("ValidationAndFreezeDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}