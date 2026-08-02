// Title: C# – Expand an existing named range with Name.RefersTo and recalculate formulas using Aspose.Cells
// Description: Demonstrates how to create a workbook, define a named range, use it in a SUM formula, change the Name.RefersTo property to a larger area, recalculate dependent formulas, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# named range | Name.RefersTo property | expand named range Aspose.Cells | recalculate formulas .NET | dynamic range Aspose.Cells | update named range programmatically | Aspose.Cells workbook.CalculateFormula
// Common Searches: how to change a named range size in Aspose.Cells C# | update Name.RefersTo and refresh formulas | expand named range and recalculate sums Aspose.Cells | Aspose.Cells replace named range example | C# code to modify named range and recalc formulas
// Developer Intent: Modify an existing named range to cover a larger cell block and refresh all formulas that reference it.
// Use Cases: Add new rows to a data table and automatically extend the named range used in totals and charts. | Adjust a report's data source before export so that all entries are included in calculations. | Programmatically synchronize a named range with a dynamic dataset for downstream processing.
// AI Prompts: Show C# code that updates the RefersTo property of a Name object to a larger range and recalculates all dependent formulas in Aspose.Cells. | Provide an Aspose.Cells .NET example that expands a named range from A1:A3 to A1:A10 and updates SUM formulas accordingly. | Explain how to verify that formulas using a renamed range reflect the new range after calling Workbook.CalculateFormula.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, define a named range, use it in a SUM formula, change the Name.RefersTo property to a larger area, recalculate dependent formulas, and save the file with Aspose.Cells for .NET.
class ReplaceNamedRange
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Populate some data in column A (A1:A5)
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1); // Values 1,2,3,4,5
        }

        // Create an initial named range that refers to A1:A3
        int nameIndex = workbook.Worksheets.Names.Add("MyRange");
        Name namedRange = workbook.Worksheets.Names[nameIndex];
        namedRange.RefersTo = "=Sheet1!$A$1:$A$3"; // original smaller area

        // Use the named range in a formula
        sheet.Cells["B1"].Formula = "=SUM(MyRange)";

        // Calculate formulas (initial sum = 1+2+3 = 6)
        workbook.CalculateFormula();
        Console.WriteLine("Initial SUM (A1:A3): " + sheet.Cells["B1"].Value);

        // Replace the existing named range with a larger area (A1:A5)
        namedRange.RefersTo = "=Sheet1!$A$1:$A$5";

        // Recalculate formulas after expanding the range (new sum = 1+2+3+4+5 = 15)
        workbook.CalculateFormula();
        Console.WriteLine("Updated SUM (A1:A5): " + sheet.Cells["B1"].Value);

        // Save the workbook (save rule)
        workbook.Save("UpdatedNamedRange.xlsx");
    }
}
