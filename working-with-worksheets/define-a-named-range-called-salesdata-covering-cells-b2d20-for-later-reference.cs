// Title: Define a Named Range "SalesData" (B2:D20) in a Workbook using Aspose.Cells for .NET (C#)
// Description: C# example that creates a new workbook with Aspose.Cells, defines the cell range B2:D20, assigns it the name "SalesData", and saves the file as SalesDataNamedRange.xlsx.
// Keywords: Aspose.Cells | C# named range | define named range | SalesData range | B2:D20 | save workbook | Aspose.Cells .NET
// Common Searches: Aspose.Cells create named range C# | how to name a cell range in Aspose.Cells | named range B2:D20 Aspose.Cells | save workbook with named range Aspose.Cells
// Developer Intent: Create and persist a named range called SalesData that covers cells B2:D20 in an Excel workbook.
// Use Cases: Reference SalesData in formulas, charts, or pivot tables without hard‑coding cell addresses. | Apply formatting, data validation, or conditional formatting to the SalesData range via its name. | Access SalesData from other worksheets or external tools that read named ranges. | Simplify source selection for pivot tables by using the SalesData named range.
// AI Prompts: Write C# code using Aspose.Cells to create a named range "SalesData" for B2:D20, apply a header style, and save the workbook. | Show how to load an existing Excel file with Aspose.Cells, retrieve the "SalesData" named range, and use it in a SUM formula. | Demonstrate adding a list‑type data validation to the "SalesData" named range with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace Example
{
    // C# example that creates a new workbook with Aspose.Cells, defines the cell range B2:D20, assigns it the name "SalesData", and saves the file as SalesDataNamedRange.xlsx.
    class DefineNamedRange
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (default name is Sheet1)
                Worksheet sheet = workbook.Worksheets[0];

                // Create a range that covers cells B2 to D20
                Aspose.Cells.Range salesRange = sheet.Cells.CreateRange("B2:D20");

                // Assign the name "SalesData" to the created range
                salesRange.Name = "SalesData";

                // Display confirmation
                Console.WriteLine($"Named range '{salesRange.Name}' created with address {salesRange.Address}");

                // Save the workbook
                string outputPath = "SalesDataNamedRange.xlsx";
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
