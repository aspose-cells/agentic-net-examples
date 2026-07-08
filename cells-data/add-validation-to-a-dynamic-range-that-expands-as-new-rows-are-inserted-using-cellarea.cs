using System;
using Aspose.Cells;

namespace AsposeCellsDynamicValidationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data in column A (A2:A5)
            for (int i = 1; i <= 4; i++) // rows are zero‑based
            {
                cells[i, 0].PutValue($"Value {i}");
            }

            // Define the initial validation area (A2:A5)
            CellArea initialArea = CellArea.CreateCellArea(1, 0, 4, 0); // rows 1‑4, column 0 (A)

            // Add a validation to the worksheet for the initial area
            int validationIndex = worksheet.Validations.Add(initialArea);
            Validation validation = worksheet.Validations[validationIndex];
            validation.Type = ValidationType.List;
            validation.Formula1 = "Option1,Option2,Option3";
            validation.InCellDropDown = true;
            validation.ShowInput = true;
            validation.ShowError = true;

            // Insert a new row at position 2 (zero‑based index 1) shifting cells down
            // This will push existing rows down, creating a new empty row at A2
            CellArea insertArea = CellArea.CreateCellArea(1, 0, 1, 0); // single cell A2
            cells.InsertRange(insertArea, 1, ShiftType.Down, true);

            // After insertion, add the newly created row (A2) to the validation range
            CellArea newRowArea = CellArea.CreateCellArea(1, 0, 1, 0);
            // No need to check intersection or edge because we know the area is unique
            validation.AddArea(newRowArea, false, false);

            // Save the workbook
            workbook.Save("DynamicValidationDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}