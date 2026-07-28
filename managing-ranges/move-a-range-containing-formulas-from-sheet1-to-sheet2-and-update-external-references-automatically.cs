// Title: Move a range with formulas from Sheet1 to Sheet2 and auto‑adjust references using Aspose.Cells for .NET
// Description: C# example that creates a workbook, fills A1:B2 on Sheet1 with values and formulas, adds Sheet2, copies the range to C3:D4 on the new sheet, automatically updates relative and external references, recalculates all formulas with CalculateFormula, and saves the file as MovedRange.xlsx.
// Keywords: Aspose.Cells | C# copy range with formulas | move range between worksheets | update formula references | CalculateFormula | Excel automation .NET | range.CopyData | external reference adjustment | Aspose.Cells example | Excel workbook manipulation
// Common Searches: Aspose.Cells copy range with formulas to another sheet | How to move a range and keep formulas working in Aspose.Cells .NET | Update external references after copying cells with Aspose.Cells | Recalculate formulas after moving a range in C# | Copy range preserving relative references Aspose.Cells
// Developer Intent: Copy a range that contains formulas to another worksheet and have all references update automatically.
// Use Cases: Duplicate a calculation block from a source sheet to a summary sheet while preserving functional formulas. | Transfer a data table with computed columns to a reporting worksheet and trigger a full recalculation in the new location. | Copy a template area that includes formulas to a new sheet for batch processing, then recalculate to obtain fresh results.
// AI Prompts: Generate C# code using Aspose.Cells that moves a range with formulas from one worksheet to another and automatically adjusts relative references. | Show how to copy a range, preserve its formulas, and invoke CalculateFormula to update all references in Aspose.Cells for .NET. | Provide an example that copies a range containing formulas to a different sheet and ensures external references are updated after the copy.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// C# example that creates a workbook, fills A1:B2 on Sheet1 with values and formulas, adds Sheet2, copies the range to C3:D4 on the new sheet, automatically updates relative and external references, recalculates all formulas with CalculateFormula, and saves the file as MovedRange.xlsx.
class MoveRangeWithFormulas
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (Sheet1) and give it a name
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Fill the source range with values and formulas
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet1.Cells["B1"].Formula = "=A1*2";   // Formula referencing the same sheet
            sheet1.Cells["B2"].Formula = "=A2*2";

            // Add a destination worksheet (Sheet2)
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Define the source range (A1:B2) on Sheet1
            AsposeRange sourceRange = sheet1.Cells.CreateRange("A1", "B2");

            // Define the destination range on Sheet2 (starting at C3, same size)
            AsposeRange destinationRange = sheet2.Cells.CreateRange("C3", "D4");

            // Copy data (including formulas) from the source range to the destination range
            destinationRange.CopyData(sourceRange);

            // Recalculate formulas so that the copied formulas produce values
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("MovedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
