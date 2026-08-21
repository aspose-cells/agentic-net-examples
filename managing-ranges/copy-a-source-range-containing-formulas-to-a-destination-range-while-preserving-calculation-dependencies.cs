// Title: Copy a range with formulas while preserving dependencies using Aspose.Cells for .NET
// Description: Creates a workbook, fills column A with numbers, adds formulas in column B that reference column A, defines source (A1:B5) and destination (A7:B11) ranges, copies the range with formulas intact, recalculates all formulas, and saves the file as an XLSX document.
// Keywords: Aspose.Cells copy range formulas .NET | preserve formula references Aspose.Cells | range.Copy method Aspose.Cells | recalculate workbook after copy | C# Excel automation Aspose | US developers Aspose.Cells | European .NET Excel library
// Common Searches: how to copy a range with formulas in Aspose.Cells C# | preserve relative cell references when copying cells Aspose.Cells | recalculate formulas after copying a range Aspose.Cells | Aspose.Cells copy range example with formulas | copy Excel block with calculations using Aspose.Cells
// Developer Intent: Duplicate a block of cells that contains formulas, ensuring the copied formulas reference the new cells correctly and the workbook is recalculated.
// Use Cases: Create a scenario analysis by replicating a calculated data block to a new area of the sheet. | Generate multiple report sections from a template that includes formulas. | Refresh a summary table after copying a pre‑calculated range to a different location.
// AI Prompts: Provide C# code that copies a range with formulas using Aspose.Cells and keeps the references relative to the new location. | Show how to force a workbook recalculation after copying a formula‑rich range with Aspose.Cells for .NET. | Explain the steps and required methods to preserve formula dependencies when copying ranges in Aspose.Cells.

using Aspose.Cells;
using System;

// Creates a workbook, fills column A with numbers, adds formulas in column B that reference column A, defines source (A1:B5) and destination (A7:B11) ranges, copies the range with formulas intact, recalculates all formulas, and saves the file as an XLSX document.
class CopyRangeWithFormulas
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source range A1:A5 with numeric values
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1); // A1..A5 = 1,2,3,4,5
            }

            // Populate source range B1:B5 with formulas that depend on column A
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 1].Formula = $"A{i + 1}*2"; // B1 = A1*2, etc.
            }

            // Define source range (A1:B5) and destination range (A7:B11)
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange(0, 0, 5, 2);      // rows 0-4, cols 0-1
            Aspose.Cells.Range destinationRange = sheet.Cells.CreateRange(6, 0, 5, 2); // rows 6-10, cols 0-1

            // Copy the source range to the destination range, preserving formulas and dependencies
            destinationRange.Copy(sourceRange);

            // Recalculate formulas so that dependent cells reflect the copied data
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("CopyRangeWithFormulas.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
