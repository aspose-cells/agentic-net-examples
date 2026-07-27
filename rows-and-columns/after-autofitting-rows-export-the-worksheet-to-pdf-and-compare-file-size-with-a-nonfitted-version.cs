using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];

        // Populate sample data (10 rows, 3 columns)
        for (int row = 0; row < 10; row++)
        {
            sheet.Cells[row, 0].PutValue($"Item {row + 1}");
            sheet.Cells[row, 1].PutValue(row * 10.5);
            sheet.Cells[row, 2].PutValue(row * 2);
        }

        // Auto‑fit rows (no fitting to pages)
        sheet.AutoFitRows();

        // Save PDF without page fitting
        string pdfPathNoFit = "Report_NoFit.pdf";
        wb.Save(pdfPathNoFit, SaveFormat.Pdf);

        // Get file size of non‑fitted PDF
        long sizeNoFit = new FileInfo(pdfPathNoFit).Length;

        // Reset workbook (or clone) to apply page‑fit settings
        // Here we reuse the same workbook and adjust PageSetup
        sheet.PageSetup.FitToPagesWide = 0;   // Fit all rows on one page (columns may span multiple pages)
        sheet.PageSetup.FitToPagesTall = 0;   // Fit all columns on one page (rows may span multiple pages)

        // Auto‑fit rows again after changing PageSetup (optional but keeps consistency)
        sheet.AutoFitRows();

        // Save PDF with page fitting
        string pdfPathFit = "Report_Fit.pdf";
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // No special options needed for this example
        };
        wb.Save(pdfPathFit, pdfOptions);

        // Get file size of fitted PDF
        long sizeFit = new FileInfo(pdfPathFit).Length;

        // Output comparison
        Console.WriteLine($"PDF size without fitting: {sizeNoFit} bytes");
        Console.WriteLine($"PDF size with fitting   : {sizeFit} bytes");
        Console.WriteLine($"Size difference         : {Math.Abs(sizeNoFit - sizeFit)} bytes");
    }
} // Author: Aspose.Cells .NET example – demonstrates auto‑fit rows, page‑fit settings, PDF export, and file‑size comparison.