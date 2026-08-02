// Title: Aspose.Cells .NET: Compare PDF file size with AllColumnsInOnePagePerSheet vs default export
// Description: C# sample that creates a workbook with 100 columns and 50 rows, saves it twice—once using default PDF settings and once with PdfSaveOptions.AllColumnsInOnePagePerSheet (and OnePagePerSheet) enabled—then reads both file sizes and outputs the byte difference.
// Keywords: Aspose.Cells | .NET PDF export | AllColumnsInOnePagePerSheet | OnePagePerSheet | PDF file size | size comparison | pagination | performance
// Common Searches: Aspose.Cells AllColumnsInOnePagePerSheet file size | PDF size difference default vs AllColumnsInOnePagePerSheet | How does OnePagePerSheet affect PDF size in Aspose.Cells | measure PDF export size Aspose.Cells .NET
// Developer Intent: Understand how enabling AllColumnsInOnePagePerSheet (with OnePagePerSheet) changes the generated PDF size compared to the default export.
// Use Cases: Assess storage impact when fitting all worksheet columns onto a single PDF page. | Choose optimal PDF save options for wide reports based on actual size differences. | Automate a size‑comparison check to decide between pagination and fit‑to‑page settings.
// AI Prompts: Write a C# method that returns the percentage size reduction when saving a workbook with AllColumnsInOnePagePerSheet versus the default PDF export. | Generate code that logs a warning if the AllColumnsInOnePagePerSheet PDF is larger than the default PDF by more than 10%. | Explain how PdfSaveOptions properties AllColumnsInOnePagePerSheet and OnePagePerSheet interact to influence PDF file size and pagination.

using System;
using System.IO;
using Aspose.Cells;

// C# sample that creates a workbook with 100 columns and 50 rows, saves it twice—once using default PDF settings and once with PdfSaveOptions.AllColumnsInOnePagePerSheet (and OnePagePerSheet) enabled—then reads both file sizes and outputs the byte difference.
class ComparePdfSizes
{
    static void Main()
    {
        // Create a new workbook and add sample data with many columns
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate 100 columns and 50 rows to ensure pagination occurs
        for (int col = 0; col < 100; col++)
        {
            sheet.Cells[0, col].PutValue("Header " + (col + 1));
            for (int row = 1; row <= 50; row++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
            }
        }

        // Save PDF with default options
        string defaultPdfPath = "default.pdf";
        workbook.Save(defaultPdfPath, SaveFormat.Pdf);

        // Save PDF with AllColumnsInOnePagePerSheet enabled
        string allColumnsPdfPath = "allcolumns.pdf";
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.AllColumnsInOnePagePerSheet = true;   // Fit all columns on one page per sheet
        pdfOptions.OnePagePerSheet = true;               // Ensure content stays on a single page
        workbook.Save(allColumnsPdfPath, pdfOptions);

        // Retrieve file sizes
        long defaultSize = new FileInfo(defaultPdfPath).Length;
        long allColumnsSize = new FileInfo(allColumnsPdfPath).Length;

        // Output comparison results
        Console.WriteLine($"Default PDF size: {defaultSize} bytes");
        Console.WriteLine($"AllColumnsInOnePagePerSheet PDF size: {allColumnsSize} bytes");
        Console.WriteLine($"Size difference: {allColumnsSize - defaultSize} bytes");
    }
}
