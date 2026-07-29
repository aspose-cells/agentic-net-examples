// Title: Unhide Columns H‑J (50‑pt width) and Export to PDF with Aspose.Cells for .NET
// Description: Loads an Excel workbook, unhides columns H to J, sets each column to a 50‑point width, and saves the result directly as a PDF using Aspose.Cells.
// Keywords: Aspose.Cells unhide columns | set column width points | Excel to PDF conversion .NET | unhide multiple columns | column visibility PDF export
// Common Searches: Aspose.Cells hide columns then export to PDF | how to set column width in points with Aspose.Cells | unhide columns H J Aspose.Cells C# | export Excel worksheet to PDF after changing column visibility | C# Aspose.Cells unhide columns range
// Developer Intent: Reveal columns H‑J, assign a 50‑point width, and generate a PDF file.
// Use Cases: Prepare a printable financial statement where hidden columns must be shown with uniform width before PDF creation. | Produce customer invoices that require specific columns to be visible and consistently sized in the final PDF. | Automate batch processing of Excel reports to unhide designated columns and archive them as PDFs.
// AI Prompts: Generate C# code with Aspose.Cells to unhide columns 7‑9, set each to 50 points, and save as PDF. | Explain the UnhideColumns method parameters and how column width in points affects PDF output in Aspose.Cells. | Show an alternative way to unhide columns individually and then export the worksheet to PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsUnhideColumnsToPdf
{
    // Loads an Excel workbook, unhides columns H to J, sets each column to a 50‑point width, and saves the result directly as a PDF using Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Unhide columns H (index 7) through J (index 9) with a width of 50 points
            // Total columns to unhide = 3 (H, I, J)
            worksheet.Cells.UnhideColumns(7, 3, 50.0);

            // Save the modified workbook as PDF
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
