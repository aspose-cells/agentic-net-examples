using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Author: Aspose.Cells .NET example – generate a landscape PDF for wide worksheets
class Program
{
    static void Main()
    {
        // Load the Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set page orientation to Landscape
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Fit all columns to a single page width; height can span multiple pages
        sheet.PageSetup.SetFitToPages(1, 0);

        // Configure PDF save options to keep all columns on one page per sheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}