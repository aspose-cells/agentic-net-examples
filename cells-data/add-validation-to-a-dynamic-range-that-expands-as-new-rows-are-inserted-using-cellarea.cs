// Title: Expand List Validation Dynamically After Row Insertion Using CellArea in Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, apply a list validation to A1:A5, insert rows with ShiftType.Down, and programmatically extend the validation range to the newly added rows by adding a CellArea to the existing Validation object before saving the file.
// Keywords: Aspose.Cells | C# validation | dynamic validation range | CellArea | list drop‑down | insert rows | ShiftType.Down | extend validation | Excel validation programmatically | Aspose.Cells .NET
// Common Searches: Aspose.Cells expand validation after inserting rows | C# add CellArea to existing validation | keep drop‑down list updated when rows are added Aspose.Cells | dynamic validation range Aspose.Cells .NET | CellArea.CreateCellArea usage example
// Developer Intent: Programmatically add a new CellArea to an existing Validation so the drop‑down list continues to cover rows inserted later.
// Use Cases: Maintain a drop‑down list that automatically includes rows added by users. | Update validation ranges in generated Excel reports without recreating the validation. | Support dynamic tables where rows are inserted during runtime. | Preserve data‑entry rules while expanding worksheets programmatically.
// AI Prompts: Generate C# code that adds a CellArea to an existing list validation after inserting rows with Aspose.Cells. | Explain how Validation.AddArea works with CellArea to keep validation ranges current. | Show a step‑by‑step example of expanding a validation range dynamically in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicValidationDemo
{
    // Shows how to create a workbook, apply a list validation to A1:A5, insert rows with ShiftType.Down, and programmatically extend the validation range to the newly added rows by adding a CellArea to the existing Validation object before saving the file.
    public class Program
    {
        public static void Main()
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

            // Add a list validation to the defined area
            int validationIndex = sheet.Validations.Add(initialArea);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.List;
            validation.InCellDropDown = true;
            validation.Formula1 = "\"Option1,Option2,Option3\"";

            // Insert three new rows after the third row (i.e., between rows 3 and 4)
            // Define the range to insert: rows 3-3 (single row) in column A
            CellArea insertArea = CellArea.CreateCellArea(3, 0, 3, 0);
            // Insert 3 rows shifting cells down and update references
            cells.InsertRange(insertArea, 3, ShiftType.Down, true);

            // After insertion, expand the validation to cover the newly inserted rows (A4:A6)
            // New rows now occupy indices 3,4,5 (zero‑based)
            CellArea newRowsArea = CellArea.CreateCellArea(3, 0, 5, 0);
            // Add the new area to the existing validation (no need to check intersection for performance)
            validation.AddArea(newRowsArea, false, false);

            // Save the workbook
            workbook.Save("DynamicValidationDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
