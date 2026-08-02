// Title: C# – Create a non‑contiguous named range and apply a SUM formula with Aspose.Cells
// Description: This example shows how to build a new workbook, fill cells A1, C1 and E1, define a named range that spans these non‑adjacent cells, assign a SUM formula that references the range, calculate the result, and save the file as NamedRangeNonContiguous.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# named range | non‑contiguous named range .NET | Excel named range multiple areas | SUM formula Aspose.Cells | Workbook.CalculateFormula | GetRanges Aspose.Cells | create named range programmatically | Aspose.Cells Excel automation
// Common Searches: how to create a named range with separate cells using Aspose.Cells | Aspose.Cells C# sum formula referencing non‑adjacent cells | retrieve areas of a multi‑area named range in Aspose.Cells | set formula for a cell that uses a non‑contiguous named range | save workbook after calculating formulas with Aspose.Cells
// Developer Intent: Define a named range that includes non‑adjacent cells and use it in a custom formula.
// Use Cases: Summarize scattered data points by grouping them into a single named range and applying SUM, AVERAGE, or MAX. | Provide a reusable data source for charts or pivot tables that require values from non‑contiguous cells. | Simplify complex formulas by referencing a multi‑area named range instead of listing each cell individually.
// AI Prompts: Write C# code with Aspose.Cells that creates a named range for cells A1, C1, and E1 and sets a SUM formula in B1. | Show how to enumerate all areas of a non‑contiguous named range using Aspose.Cells. | Demonstrate using the same named range in an AVERAGE formula with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeDemo
{
    // This example shows how to build a new workbook, fill cells A1, C1 and E1, define a named range that spans these non‑adjacent cells, assign a SUM formula that references the range, calculate the result, and save the file as NamedRangeNonContiguous.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Populate some sample data in non‑contiguous cells
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["C1"].PutValue(20);
                sheet.Cells["E1"].PutValue(30);

                // Add a named range that refers to the non‑contiguous cells
                // The RefersTo string uses commas to separate the areas
                int nameIndex = sheet.Workbook.Worksheets.Names.Add("MyNonContig");
                Name myName = sheet.Workbook.Worksheets.Names[nameIndex];
                myName.RefersTo = "=Sheet1!$A$1,$C$1,$E$1";

                // Verify the named range by retrieving all its areas
                // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
                Aspose.Cells.Range[] areas = myName.GetRanges();
                Console.WriteLine("Named range consists of the following areas:");
                foreach (Aspose.Cells.Range area in areas)
                {
                    Console.WriteLine($" - {area.RefersTo}");
                }

                // Assign a custom formula that uses the named range (sum of the three cells)
                sheet.Cells["B1"].Formula = "=SUM(MyNonContig)";

                // Calculate formulas so that B1 gets the result
                workbook.CalculateFormula();

                // Output the result of the custom formula
                Console.WriteLine($"Result of SUM(MyNonContig) in B1: {sheet.Cells["B1"].Value}");

                // Save the workbook (lifecycle save)
                string outputPath = "NamedRangeNonContiguous.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
