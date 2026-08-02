// Title: C# – Update a Named Range and Recalculate Formulas with Workbook.CalculateFormula in Aspose.Cells
// Description: Demonstrates how to create a workbook, define a named range, use it in a SUM formula, extend the range, and call Workbook.CalculateFormula to refresh all dependent cells in a .NET application using Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | named range | Workbook.CalculateFormula | recalculate formulas | update named range | Excel automation | SUM formula | programmatic range extension
// Common Searches: Aspose.Cells recalculate formulas after named range change | C# update named range and refresh dependent cells | Workbook.CalculateFormula example with named ranges | extend named range programmatically Aspose.Cells | how to refresh formulas when named range is modified
// Developer Intent: Refresh all formulas that reference a named range after the range definition is modified.
// Use Cases: Generate a financial worksheet where monthly values are added to a named range and totals must update automatically. | Import data batches into an Excel file, expand the named range, and ensure all calculations reflect the new rows. | Create a reporting tool that adjusts named ranges on the fly and instantly recalculates dependent charts and summaries.
// AI Prompts: Write C# code using Aspose.Cells to modify an existing named range and automatically recalculate all formulas that reference it. | Show how to add rows to a named range, update its RefersTo property, and invoke Workbook.CalculateFormula to refresh dependent cells in a .NET project.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, define a named range, use it in a SUM formula, extend the range, and call Workbook.CalculateFormula to refresh all dependent cells in a .NET application using Aspose.Cells.
class UpdateNamedRangeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Access the first worksheet
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Data";

        // Populate some sample data
        ws.Cells["A1"].PutValue(10);
        ws.Cells["A2"].PutValue(20);
        ws.Cells["A3"].PutValue(30);

        // Create a named range that refers to A1:A3
        int nameIdx = wb.Worksheets.Names.Add("MyRange");
        Name myRange = wb.Worksheets.Names[nameIdx];
        myRange.RefersTo = "=Sheet1!$A$1:$A$3";

        // Use the named range in a formula
        ws.Cells["B1"].Formula = "=SUM(MyRange)";

        // Calculate formulas so B1 reflects the sum of the range
        wb.CalculateFormula();
        Console.WriteLine("Initial sum (A1:A3): " + ws.Cells["B1"].Value);

        // Extend the data range by adding a new value
        ws.Cells["A4"].PutValue(40);

        // Update the named range to include the new cell
        myRange.RefersTo = "=Sheet1!$A$1:$A$4";

        // Propagate the change by recalculating formulas
        wb.CalculateFormula();
        Console.WriteLine("Updated sum (A1:A4): " + ws.Cells["B1"].Value);

        // Save the workbook (optional)
        wb.Save("UpdatedNamedRange.xlsx");
    }
}
