// Title: Recalculate all workbook formulas, including dynamic array formulas, after updating cell values with Aspose.Cells for .NET (C#)
// Common Searches: how to recalculate dynamic array formulas after changing cell values using Aspose.Cells C# | Aspose.Cells C# refresh all formulas in a workbook after modifying cells | C# workbook.CalculateFormula to update formulas across multiple worksheets | update cell A1 and B2 then recalculate Excel formulas with Aspose.Cells .NET | recompute Excel formulas after cell edit using Aspose.Cells for .NET
// Tags: full workbook formula recalculation Aspose.Cells | dynamic array formula refresh .NET | cell edit triggers workbook.CalculateFormula | cross-sheet formula update Aspose.Cells | C# modify cells then recalc Excel workbook

using Aspose.Cells;
using System;

// The example loads an existing Excel file, changes the value of A1 in the first worksheet and B2 in the second worksheet (if present), calls workbook.CalculateFormula() to recompute all formulas—including dynamic array formulas—across every sheet, and then saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Modify cell values as needed
        // Example: set A1 in the first worksheet to a numeric value
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Cells["A1"].PutValue(10);

        // Example: set B2 in the second worksheet to a string value
        if (workbook.Worksheets.Count > 1)
        {
            Worksheet sheet2 = workbook.Worksheets[1];
            sheet2.Cells["B2"].PutValue("Test");
        }

        // Refresh all formulas (including dynamic array formulas) across all sheets
        workbook.CalculateFormula();

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}
