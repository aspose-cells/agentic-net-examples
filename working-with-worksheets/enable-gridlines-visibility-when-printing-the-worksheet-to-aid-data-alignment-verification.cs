// Title: Print Gridlines in an Excel Worksheet Using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add sample cells, display gridlines on screen, enable gridline printing via PageSetup.PrintGridlines, and save the file as GridlinesPrinted.xlsx.
// Keywords: Aspose.Cells | C# | .NET | PrintGridlines | worksheet gridlines | Excel printing | PageSetup | IsGridlinesVisible
// Common Searches: Aspose.Cells print gridlines .NET | Enable gridlines when printing Excel with Aspose | PageSetup.PrintGridlines example | Show gridlines in printed workbook Aspose.Cells | C# Aspose.Cells gridline visibility
// Developer Intent: Add gridlines to the printed output of an Excel worksheet.
// Use Cases: Display gridlines while editing a worksheet for visual reference. | Include gridlines in PDF, XPS, or printed Excel output. | Prepare a report where data alignment is verified by visible gridlines.
// AI Prompts: Write C# code that sets PrintGridlines for all worksheets in a workbook and exports to PDF. | Compare IsGridlinesVisible and PageSetup.PrintGridlines in Aspose.Cells and show when to use each. | Provide a tutorial for printing an Excel sheet with gridlines to XPS using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsGridlinesPrintDemo
{
    // Demonstrates how to create a workbook, add sample cells, display gridlines on screen, enable gridline printing via PageSetup.PrintGridlines, and save the file as GridlinesPrinted.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data so the gridlines are visible in the output
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["A2"].PutValue("Data 1");
            worksheet.Cells["B2"].PutValue("Data 2");

            // Ensure gridlines are shown on the screen (optional, does not affect printing)
            worksheet.IsGridlinesVisible = true;

            // Enable printing of gridlines
            worksheet.PageSetup.PrintGridlines = true;

            // Save the workbook; the printed version will include gridlines
            workbook.Save("GridlinesPrinted.xlsx");
        }
    }
}
