// Title: Create a non‑contiguous named range and apply a SUM formula with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to build a new workbook, fill cells A1, B3 and D5, define a named range that spans these non‑adjacent cells using the RefersTo property, assign a SUM formula that references the range, calculate the result, and save the file as NonContiguousNamedRange.xlsx.
// Keywords: Aspose.Cells C# | .NET Excel automation | named range non‑contiguous | RefersTo multiple areas | custom formula Aspose.Cells | SUM named range | create named range programmatically | Excel workbook Aspose | non adjacent cells formula
// Common Searches: Aspose.Cells create named range with separate cells | How to use RefersTo for non‑contiguous ranges in C# | Apply SUM to a multi‑area named range using Aspose.Cells | Define named range that includes A1, B3, D5 in .NET | Calculate formulas after adding a named range in Aspose.Cells
// Developer Intent: Define a named range that points to non‑adjacent cells and use it in a custom formula.
// Use Cases: Aggregate scattered data points with a single SUM expression. | Provide a reusable range for charts or pivot tables that pull values from non‑contiguous cells. | Simplify complex worksheets by grouping unrelated cells under one name for formula reuse.
// AI Prompts: Generate C# code with Aspose.Cells that creates a named range covering cells A1, C2, and E3 and sets an AVERAGE formula using that range. | Show how to extend an existing named range's RefersTo property to include an additional non‑contiguous cell and refresh dependent formulas. | Write a script that creates multiple non‑contiguous named ranges and applies different custom formulas (SUM, MAX, MIN) to each.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to build a new workbook, fill cells A1, B3 and D5, define a named range that spans these non‑adjacent cells using the RefersTo property, assign a SUM formula that references the range, calculate the result, and save the file as NonContiguousNamedRange.xlsx.
    public class NonContiguousNamedRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and give it a name
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Populate some sample data in non‑contiguous cells
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["D5"].PutValue(30);

                // Add a named range that refers to the non‑contiguous cells
                int nameIndex = workbook.Worksheets.Names.Add("MyNonContig");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                // RefersTo can contain multiple areas separated by commas
                namedRange.RefersTo = "=Sheet1!$A$1,$B$3,$D$5";

                // Assign a custom formula that uses the named range (e.g., sum of the cells)
                sheet.Cells["F1"].Formula = "=SUM(MyNonContig)";

                // Calculate formulas so that the result appears in F1
                workbook.CalculateFormula();

                // Save the workbook
                workbook.Save("NonContiguousNamedRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                NonContiguousNamedRangeDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
