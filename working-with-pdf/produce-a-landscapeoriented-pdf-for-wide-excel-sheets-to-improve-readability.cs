// Title: Generate a landscape-oriented PDF from a wide Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets each worksheet's PageSetup.Orientation to Landscape, configures FitToPagesWide = 1 and FitToPagesTall = 0, and saves the workbook as a PDF. | Show how to programmatically export a multi‑column Excel sheet to a single‑page‑wide landscape PDF using Aspose.Cells PageSetup settings in a .NET application.
// Common Searches: Aspose.Cells C# export wide Excel sheet to landscape PDF with fit to width | How to set page orientation to landscape when converting Excel to PDF using Aspose.Cells .NET | C# Aspose.Cells fit worksheet to one PDF page width landscape | Convert Excel workbook to PDF landscape orientation without cutting columns Aspose.Cells | Save Excel as PDF landscape and keep all columns visible Aspose.Cells C#
// Tags: Aspose.Cells PDF landscape export | PageSetup orientation landscape Aspose.Cells | FitToPagesWide single page PDF Aspose.Cells | C# convert wide Excel to PDF | Excel to PDF fit width landscape

using System;
using Aspose.Cells;

// // Loads input.xlsx, sets each worksheet to landscape orientation, fits width to one PDF page (FitToPagesWide = 1, FitToPagesTall = 0), and saves as output.pdf.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure each worksheet to use landscape orientation
        // This improves readability for wide sheets when exported to PDF
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Set page orientation to landscape
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Optional: fit the sheet width to a single PDF page while allowing unlimited length
            sheet.PageSetup.FitToPagesWide = 1;   // fit to one page wide
            sheet.PageSetup.FitToPagesTall = 0;   // no restriction on page height
        }

        // Save the workbook as a PDF file; the orientation settings are applied automatically
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
