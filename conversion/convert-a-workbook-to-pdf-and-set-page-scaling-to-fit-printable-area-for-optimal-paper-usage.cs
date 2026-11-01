// Title: Convert an Excel workbook to PDF with page scaling to fit the printable area using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to PDF with FitToPagesWide set to 1 | How to fit worksheet content to one page width when saving as PDF in Aspose.Cells | C# code to set page setup for PDF conversion to use printable area in Aspose.Cells | Scale Excel sheets to printable area before PDF generation with Aspose.Cells .NET | Adjust page scaling for PDF output of multiple worksheets using Aspose.Cells
// Tags: Aspose.Cells PDF export page scaling | FitToPagesWide property Excel to PDF | set printable area before PDF save Aspose.Cells | C# workbook conversion to PDF with page setup | scale worksheets to single page width Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook, configures each worksheet's PageSetup to fit one page wide (FitToPagesWide = 1) while allowing automatic height (FitToPagesTall = 0), and then saves the workbook as a PDF, ensuring the content uses the printable area efficiently.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure each worksheet to fit the printable area on a single page width
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;
            // Fit to 1 page wide; 0 for tall means auto‑fit to the required number of pages
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 0;
        }

        // Save the workbook as a PDF with the scaling applied
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
