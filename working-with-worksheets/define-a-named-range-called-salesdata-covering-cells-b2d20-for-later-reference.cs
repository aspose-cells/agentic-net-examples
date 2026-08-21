// Title: C# Example: Define and Save a Named Range "SalesData" (B2:D20) with Aspose.Cells
// Description: Demonstrates how to create a new Workbook, access the first worksheet, build a range covering cells B2 through D20, assign the name "SalesData" to that range, and persist the file as NamedRangeDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# named range | create named range B2:D20 | SalesData range Aspose | C# workbook save with named range | .NET Excel range naming example | Aspose.Cells range API
// Common Searches: Aspose.Cells how to add a named range in C# | define SalesData range B2 to D20 Aspose.Cells | C# save Excel file with named range using Aspose | Aspose.Cells .NET create named range programmatically
// Developer Intent: Add a named range called SalesData that spans B2:D20 and write the workbook to disk.
// Use Cases: Reference SalesData in formulas or chart data sources without hard‑coding cell addresses. | Export the workbook to other platforms that recognize named ranges, preserving the logical grouping of sales data. | Iterate over the SalesData range programmatically for custom validation, transformation, or reporting.
// AI Prompts: Write C# code with Aspose.Cells that creates a named range 'SalesData' for B2:D20 and then uses it in a SUM formula. | Show how to open an existing workbook, retrieve the 'SalesData' named range, and loop through its cells in C#. | Explain the steps to rename or delete a named range in an Aspose.Cells workbook using .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create a new Workbook, access the first worksheet, build a range covering cells B2 through D20, assign the name "SalesData" to that range, and persist the file as NamedRangeDemo.xlsx using Aspose.Cells for .NET.
class DefineNamedRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (default name is "Sheet1")
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a range that covers cells B2:D20 using the Aspose.Cells Range class
            Aspose.Cells.Range salesRange = worksheet.Cells.CreateRange("B2", "D20");

            // Assign a name to the range for later reference
            salesRange.Name = "SalesData";

            // Save the workbook (the file will contain the named range)
            workbook.Save("NamedRangeDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
