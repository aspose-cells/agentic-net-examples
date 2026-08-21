// Title: Aspose.Cells .NET: Update a Named Range and Recalculate Dependent Formulas with Workbook.CalculateFormula
// Description: Demonstrates how to create a workbook, define a named range, use it in a SUM formula, expand the range, modify the Name.RefersTo property, invoke Workbook.CalculateFormula to refresh all dependent calculations, and finally save the file.
// Keywords: Aspose.Cells | C# named range update | Workbook.CalculateFormula | recalculate formulas after name change | expand named range sum | Aspose.Cells .NET example | Name.RefersTo property
// Common Searches: Aspose.Cells recalculate after named range modification | How to refresh formulas when Name.RefersTo changes | C# update named range and recalc with Aspose.Cells | Workbook.CalculateFormula usage example | Expand named range in Aspose.Cells .NET
// Developer Intent: Refresh all formulas that reference a named range after the range definition has been altered.
// Use Cases: Add new cells to an existing named range and automatically update SUM or other formulas that use the name. | Programmatically change the RefersTo expression of a Name object and ensure calculation results stay current. | Create, modify, and persist a workbook where formula results depend on dynamic named ranges.
// AI Prompts: Write C# code that changes a Name.RefersTo value in an Aspose.Cells workbook and calls Workbook.CalculateFormula to update dependent cells. | Show an example of expanding a named range from A1:A3 to A1:A4 and recalculating a SUM formula using Aspose.Cells for .NET. | Explain the effect of Workbook.CalculateFormula after modifying a named range in Aspose.Cells and how it impacts cached formula results.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeUpdateDemo
{
    // Demonstrates how to create a workbook, define a named range, use it in a SUM formula, expand the range, modify the Name.RefersTo property, invoke Workbook.CalculateFormula to refresh all dependent calculations, and finally save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];

            // Populate some sample data that will be referenced by a named range
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Create a named range "MyRange" that refers to A1:A3
            int nameIndex = wb.Worksheets.Names.Add("MyRange");
            Name myRange = wb.Worksheets.Names[nameIndex];
            myRange.RefersTo = "=Sheet1!$A$1:$A$3";

            // Use the named range in a formula (e.g., sum of the range)
            sheet.Cells["B1"].Formula = "=SUM(MyRange)";

            // Calculate formulas so that B1 reflects the initial sum (30)
            wb.CalculateFormula();

            Console.WriteLine("Initial sum (B1): " + sheet.Cells["B1"].Value); // Expected: 60

            // Update the named range to include an additional cell (A4)
            sheet.Cells["A4"].PutValue(40);               // Add new data
            myRange.RefersTo = "=Sheet1!$A$1:$A$4";       // Change the RefersTo formula

            // Propagate the change by recalculating formulas (required after named range update)
            wb.CalculateFormula();

            Console.WriteLine("Updated sum (B1) after expanding named range: " + sheet.Cells["B1"].Value); // Expected: 100

            // Save the workbook (lifecycle rule: save)
            wb.Save("NamedRangeUpdateResult.xlsx");
        }
    }
}
