// Title: Cut and paste a formula range with Aspose.Cells for .NET while preserving references
// Description: Demonstrates how to cut a range that contains formulas (A1:B3) and insert it at a new location (C1) using Aspose.Cells' InsertCutCells with ShiftType.Down, ensuring all formula references are automatically updated before saving the workbook.
// Keywords: Aspose.Cells cut range | move cells with formulas .NET | InsertCutCells example | preserve formula references | ShiftType.Down Aspose | C# Excel manipulation | cut and paste range Aspose.Cells
// Common Searches: Aspose.Cells cut range with formulas | how to preserve formula references when moving cells in .NET | InsertCutCells ShiftType.Down usage | C# cut and paste Excel range Aspose | move calculated block without breaking formulas
// Developer Intent: Relocate a block of cells that includes formulas, updating all references automatically.
// Use Cases: Rearrange a calculated table from columns A‑B to C‑D at runtime without breaking dependent formulas. | Generate a dynamic report by cutting a pre‑formatted formula section and inserting it into a designated report area. | Programmatically shift worksheet layout downward while keeping all calculations intact.
// AI Prompts: Write C# code that cuts range A1:B3 containing formulas and inserts it at C1 using Aspose.Cells, preserving formula references. | Explain how InsertCutCells with ShiftType.Down updates relative formulas when a range is moved in Aspose.Cells for .NET. | Provide a step‑by‑step example of cutting a formula range and inserting it into a new location while shifting existing cells down.

using System;
using Aspose.Cells;

namespace AsposeCellsCutPasteDemo
{
    // Demonstrates how to cut a range that contains formulas (A1:B3) and insert it at a new location (C1) using Aspose.Cells' InsertCutCells with ShiftType.Down, ensuring all formula references are automatically updated before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Fill the source range with values
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["A3"].PutValue(30);

                // Add formulas that reference the values in column A
                cells["B1"].Formula = "=A1*2";
                cells["B2"].Formula = "=A2*2";
                cells["B3"].Formula = "=A3*2";

                // Define the range to cut (A1:B3)
                Aspose.Cells.Range cutRange = cells.CreateRange("A1:B3");

                // Insert the cut range at a new location (C1) shifting cells down.
                // Row index = 0 (first row), Column index = 2 (column C)
                cells.InsertCutCells(cutRange, 0, 2, ShiftType.Down);

                // Save the workbook with the cut‑and‑pasted data
                string outputPath = "CutPasteDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
