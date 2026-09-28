// Title: Measure performance of exporting a 500‑sheet workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook with 500 worksheets, fills each with sample data, configures PdfSaveOptions for one page per sheet, and records the time taken to save the workbook to a PDF stream using Aspose.Cells. | Write a C# benchmark that evaluates the scalability of Aspose.Cells PDF conversion by timing the Save operation for a large multi‑sheet workbook.
// Common Searches: how long does Aspose.Cells take to export a 500 sheet workbook to PDF in C# | performance test for Aspose.Cells PDFSaveOptions with many worksheets | benchmark Aspose.Cells PDF conversion speed for large workbooks | measure scalability of Aspose.Cells when saving multi‑sheet workbooks to PDF | C# code to time Aspose.Cells workbook.Save for 500 worksheets
// Tags: Aspose.Cells PDFSaveOptions performance | export large workbook to PDF .NET | benchmark Aspose.Cells multi‑sheet PDF conversion | C# timing Aspose.Cells workbook.Save | scalability testing Aspose.Cells PDF export

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The program builds a workbook containing 500 worksheets, adds header and sample rows to each sheet, configures PdfSaveOptions to place each worksheet on a separate page and fit all columns, then saves the workbook to a PDF in a memory stream while measuring and printing the elapsed time.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and clear the default worksheet
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();

            // Add 500 worksheets with sample data
            for (int i = 0; i < 500; i++)
            {
                Worksheet sheet = workbook.Worksheets.Add("Sheet" + (i + 1));

                // Header row
                sheet.Cells["A1"].PutValue("Row");
                sheet.Cells["B1"].PutValue("Value");

                // Populate rows 2..100
                for (int row = 2; row <= 100; row++)
                {
                    sheet.Cells[row, 0].PutValue(row - 1);
                    sheet.Cells[row, 1].PutValue(Math.Sin(row));
                }
            }

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export each worksheet on a separate page
                OnePagePerSheet = true,
                // Fit all columns of a sheet onto one page
                AllColumnsInOnePagePerSheet = true
            };

            // Benchmark the export operation
            Stopwatch sw = Stopwatch.StartNew();

            // Save to a memory stream to avoid file I/O overhead
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, pdfOptions);
            }

            sw.Stop();

            Console.WriteLine($"Export of 500 worksheets completed in {sw.Elapsed.TotalSeconds} seconds.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
