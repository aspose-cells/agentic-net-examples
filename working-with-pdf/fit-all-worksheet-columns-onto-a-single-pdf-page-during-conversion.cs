// Title: How to fit all worksheet columns onto a single PDF page using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# fit worksheet columns to one PDF page width | Set FitToPagesWide = 1 for PDF export with Aspose.Cells | Convert Excel to PDF with all columns on a single page using Aspose.Cells | C# adjust worksheet PageSetup before saving as PDF with Aspose.Cells | Export Excel workbook to PDF with columns scaled to one page in .NET
// Tags: FitToPagesWide column scaling PDF export | Aspose.Cells page setup PDF conversion | C# fit columns to single PDF page | Excel to PDF column fit Aspose.Cells | Worksheet page setup for PDF in .NET

using Aspose.Cells;
using System;

// The example loads an Excel workbook, sets each worksheet's PageSetup to FitToPagesWide = 1 and FitToPagesTall = 0 so all columns fit within one PDF page width, and then saves the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure each worksheet to fit all columns on a single PDF page width
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Fit all columns to one page wide; height is unrestricted (0 = auto)
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;
        }

        // Save the workbook as a PDF file with the applied page settings
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
