// Title: Print Gridlines and Row/Column Headings in Aspose.Cells for .NET
// Description: This C# sample builds a workbook, fills cells A1:B3, turns on on‑screen grid lines, activates PageSetup.PrintGridlines and PageSetup.PrintHeadings, sets a print area, and exports the sheet to PDF so the output contains the grid and the row/column labels.
// Keywords: Aspose.Cells | C# | PrintGridlines | PrintHeadings | PDF export | PageSetup | grid lines | row headings | column headings | print area | worksheet printing
// Common Searches: Aspose.Cells enable grid lines in PDF | How to print row and column headings with Aspose.Cells | Set print area and include headings in Aspose.Cells C# | Export Excel to PDF with grid lines using Aspose | Print worksheet with headings Aspose.Cells .NET
// Developer Intent: Create a PDF that mirrors the Excel view by showing both the cell borders and the row/column identifiers.
// Use Cases: Produce a printable product catalog where the table layout, including borders and labels, matches the on‑screen Excel sheet. | Generate an invoice PDF that retains the exact grid structure and header rows for audit‑ready documentation. | Export analytical data worksheets to PDF while preserving visual cues such as grid lines and axis headings for reports.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet to PDF showing grid lines and row/column headings. | Show how to configure PrintGridlines, PrintHeadings, and a custom PrintArea before saving as PDF in Aspose.Cells. | Explain the steps to make grid lines visible on screen and ensure they appear in the printed PDF together with headings using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# sample builds a workbook, fills cells A1:B3, turns on on‑screen grid lines, activates PageSetup.PrintGridlines and PageSetup.PrintHeadings, sets a print area, and exports the sheet to PDF so the output contains the grid and the row/column labels.
    public class PrintGridlinesAndHeadingsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(2.5);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(1.8);

            // Make gridlines visible on screen (optional)
            worksheet.IsGridlinesVisible = true;

            // Enable printing of gridlines and row/column headings
            worksheet.PageSetup.PrintGridlines = true;
            worksheet.PageSetup.PrintHeadings = true;

            // Define the print area
            worksheet.PageSetup.PrintArea = "A1:B3";

            // Save the workbook (PDF demonstrates the printed result)
            workbook.Save("PrintGridlinesAndHeadings.pdf");
        }
    }
}
