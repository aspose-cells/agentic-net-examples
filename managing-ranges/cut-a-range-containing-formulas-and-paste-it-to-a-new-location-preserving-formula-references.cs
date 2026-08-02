// Title: Cut and paste a formula range while preserving references with Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, place values and formulas in A1:C2, cut that range, and insert it at A5 using Aspose.Cells' InsertCutCells with ShiftType.Down. Formula links are automatically updated and the file is saved as CutPastePreserveFormulas.xlsx.
// Keywords: Aspose.Cells | C# | .NET | cut range | paste range | preserve formulas | InsertCutCells | ShiftType.Down | move cells with formulas | range manipulation | Excel automation | cell shifting
// Common Searches: Aspose.Cells cut range with formulas | move cells and keep formulas .NET | InsertCutCells C# example | ShiftType.Down usage Aspose.Cells | cut and paste cells preserving references | range copy paste Aspose.Cells | adjust formula links after moving cells
// Developer Intent: Shift a block of cells that includes formulas to a new position without breaking dependent calculations.
// Use Cases: Reorder a budgeting section by cutting a calculation block and inserting it lower in the sheet while all formulas stay correct. | Create space for new entries by moving a data table down, ensuring summary formulas continue to reference the right cells. | Automate report layout changes by relocating summary rows to a designated area without manually fixing formula references.
// AI Prompts: Provide a C# example that cuts a range containing formulas and pastes it elsewhere using Aspose.Cells, keeping all references intact. | How does InsertCutCells with ShiftType.Down adjust formula links when moving a cell block in an Aspose.Cells workbook? | Show code to relocate cells A1:C2 to A5 in a .NET workbook while automatically updating any dependent formulas.

using System;
using Aspose.Cells;

namespace AsposeCellsCutPasteDemo
{
    // Shows how to build a workbook, place values and formulas in A1:C2, cut that range, and insert it at A5 using Aspose.Cells' InsertCutCells with ShiftType.Down. Formula links are automatically updated and the file is saved as CutPastePreserveFormulas.xlsx.
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

                // Populate source range with values and formulas
                cells["A1"].PutValue(10);                     // Simple value
                cells["B1"].PutValue(20);                     // Simple value
                cells["C1"].Formula = "=A1+B1";               // Formula referencing A1 and B1
                cells["A2"].Formula = "=C1*2";                // Formula referencing the formula cell

                // Define the range to cut (A1:C2)
                Aspose.Cells.Range cutRange = cells.CreateRange("A1:C2");

                // Insert the cut range at a new location (starting at row 4, column 0 i.e., A5)
                // ShiftType.Down will shift existing cells down to make space
                cells.InsertCutCells(cutRange, 4, 0, ShiftType.Down);

                // Save the workbook to verify the operation
                workbook.Save("CutPastePreserveFormulas.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
