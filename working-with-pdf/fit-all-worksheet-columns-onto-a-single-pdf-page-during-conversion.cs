// Title: Fit all worksheet columns onto a single PDF page during Excel-to-PDF conversion using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that configures each worksheet's PageSetup so that columns are scaled to one PDF page width before saving with Aspose.Cells. | Show how to set FitToPagesWide = 1 and FitToPagesTall = 0 for PDF export in Aspose.Cells .NET. | Provide a complete example that loads an .xlsx file, applies page‑setup settings to force a single‑page column layout, and saves the workbook as a PDF.
// Common Searches: Aspose.Cells C# export Excel to PDF with columns forced to one page width | How to make worksheet columns fit on a single PDF page using Aspose.Cells .NET | Set page setup for PDF conversion to fit all columns in Aspose.Cells | FitToPagesWide property usage example for PDF output in Aspose.Cells | Convert large Excel sheet to PDF with one‑page column scaling Aspose.Cells
// Tags: column scaling to one PDF page Aspose.Cells | FitToPagesWide property C# | worksheet page setup for PDF export | Aspose.Cells PDF conversion column scaling | set FitToPagesTall zero Aspose.Cells

using Aspose.Cells;
using System;

// // Loads an Excel workbook, sets each worksheet's PageSetup to FitToPagesWide = 1 and FitToPagesTall = 0 so all columns fit on one PDF page, and saves the result as a PDF file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure each worksheet to fit all columns on a single PDF page
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // One page wide, unlimited pages tall (0 means auto)
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;
        }

        // Save the workbook as PDF (replace with your desired output path)
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
