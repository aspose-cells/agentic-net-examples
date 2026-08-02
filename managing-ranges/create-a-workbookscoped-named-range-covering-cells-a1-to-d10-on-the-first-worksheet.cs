// Title: Aspose.Cells .NET: Create a Workbook‑Scoped Named Range for A1:D10
// Description: Demonstrates how to create a new Workbook, define a range covering cells A1 to D10 on the first worksheet, assign a workbook‑scoped name (global) "MyRange", and save the file as WorkbookScopedNamedRange.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells named range C# | workbook scoped range .NET | global named range Aspose.Cells | create range A1 D10 | Aspose.Cells C# example
// Common Searches: Aspose.Cells create workbook scoped named range C# | global named range A1:D10 Aspose.Cells | how to set named range without sheet prefix Aspose.Cells | Aspose.Cells define range programmatically .NET | C# Aspose.Cells named range example
// Developer Intent: Create a workbook‑scoped (global) named range that spans A1:D10 and persist it in a workbook file.
// Use Cases: Use the global name "MyRange" in formulas on any worksheet without a sheet qualifier. | Apply formatting, data validation, or calculations to the same cell block across multiple sheets. | Export or import the named range data for integration with reporting or analytics tools.
// AI Prompts: Generate C# code to retrieve the workbook‑scoped named range "MyRange" and loop through its cells. | Show how to create a worksheet‑scoped named range and compare its scope to a workbook‑scoped range. | Provide a method that lists all workbook‑scoped named ranges in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeDemo
{
    // Demonstrates how to create a new Workbook, define a range covering cells A1 to D10 on the first worksheet, assign a workbook‑scoped name (global) "MyRange", and save the file as WorkbookScopedNamedRange.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a range that covers cells A1 to D10
                AsposeRange range = worksheet.Cells.CreateRange("A1", "D10");

                // Assign a workbook‑scoped name to the range.
                // Setting the Name property without a sheet prefix makes it global.
                range.Name = "MyRange";

                // Save the workbook to a file
                workbook.Save("WorkbookScopedNamedRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
