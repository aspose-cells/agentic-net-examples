// Title: How to remove worksheet printer settings, set A2 paper size, and export to PDF with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# clear printer settings for all worksheets before PDF export | set A2 paper size using Aspose.Cells PageSetup in C# | remove printer configuration and save workbook as PDF with Aspose.Cells .NET | C# Aspose.Cells change page setup to A2 and generate PDF
// Tags: clear worksheet printer settings Aspose.Cells | A2 paper size PageSetup Aspose.Cells | export workbook to PDF Aspose.Cells | remove PrinterSettings Aspose.Cells .NET | modify page setup for PDF conversion Aspose.Cells

using System;
using Aspose.Cells;
using System.Drawing.Printing;

// The code loads a workbook, clears PrinterSettings and sets PaperSize to A2 for each worksheet, then saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets to modify page settings
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;

            // Remove any existing printer settings
            pageSetup.PrinterSettings = null;

            // Set the paper size to A2
            pageSetup.PaperSize = PaperSizeType.PaperA2;
        }

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
