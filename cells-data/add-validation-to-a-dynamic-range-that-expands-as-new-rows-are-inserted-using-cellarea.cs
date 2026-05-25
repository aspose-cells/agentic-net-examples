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
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data in column A (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue($"Item {i + 1}");
            }

            // Define the initial validation area (A1:A5)
            CellArea initialArea = CellArea.CreateCellArea(0, 0, 4, 0); // rows 0-4, column 0

            // Add a list‑type validation to the initial area
            int validationIndex = sheet.Validations.Add(initialArea);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.List;
            validation.InCellDropDown = true;
            validation.Formula1 = "\"Option1,Option2,Option3\"";

            // Insert a new row at position 2 (zero‑based index 1) and shift cells down.
            // The 'updateReference' flag ensures that references (including validation ranges) are adjusted.
            CellArea insertArea = CellArea.CreateCellArea(1, 0, 1, 0); // single cell A2
            cells.InsertRange(insertArea, 1, ShiftType.Down, true);

            // After insertion the original validation now covers A1:A6 automatically because of updateReference.
            // To demonstrate dynamic expansion, add an additional validation area for the newly inserted row.
            // (This step is optional; the range is already expanded, but we show explicit addition.)
            CellArea newRowArea = CellArea.CreateCellArea(1, 0, 1, 0); // the inserted row A2
            validation.AddArea(newRowArea, false, false);

            // Save the workbook
            workbook.Save("DynamicValidationDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}