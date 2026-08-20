// Title: C# – Set 0.5‑inch margins and fit worksheet to a single page with Aspose.Cells for .NET
// Description: Shows how to create a workbook, apply 0.5 inch left, right, top and bottom margins via PageSetup, scale the sheet with SetFitToPages(1,1) so it prints on one page, and save the result as an XLSX file.
// Keywords: Aspose.Cells C# set margins | 0.5 inch page margins Aspose.Cells | fit worksheet to one page .NET | PageSetup SetFitToPages | scale worksheet to single page | custom page margins Aspose.Cells | print layout Aspose.Cells | C# Excel margin settings | Aspose.Cells print scaling
// Common Searches: Aspose.Cells set 0.5 inch margins C# | How to fit Excel sheet to one page using Aspose.Cells | PageSetup margins inches Aspose.Cells .NET | Scale worksheet to single printed page C# | SetFitToPages example Aspose.Cells
// Developer Intent: Apply half‑inch margins on every side of a worksheet and then scale it to print on a single page.
// Use Cases: Generating printable reports that must stay within half‑inch margins and occupy one page. | Creating invoices that automatically adjust to a single‑page layout regardless of row count. | Exporting data tables to PDF with consistent margins and a fit‑to‑page format.
// AI Prompts: Provide C# code that sets 0.5‑inch margins on all sides and fits a worksheet to one page using Aspose.Cells. | Show an example of configuring PageSetup margins in inches and applying SetFitToPages(1,1) for a workbook. | Explain how to combine custom margin settings with scaling options to ensure a worksheet prints on a single page in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsMarginAndFitDemo
{
    // Shows how to create a workbook, apply 0.5 inch left, right, top and bottom margins via PageSetup, scale the sheet with SetFitToPages(1,1) so it prints on one page, and save the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set custom margins of 0.5 inches on all sides
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.LeftMarginInch = 0.5;
            pageSetup.RightMarginInch = 0.5;
            pageSetup.TopMarginInch = 0.5;
            pageSetup.BottomMarginInch = 0.5;

            // Scale the worksheet to fit on one page (both width and height)
            pageSetup.SetFitToPages(1, 1);
            // Alternatively, you could use:
            // pageSetup.FitToPagesWide = 1;
            // pageSetup.FitToPagesTall = 1;

            // Save the workbook
            workbook.Save("output.xlsx");
        }
    }
}
