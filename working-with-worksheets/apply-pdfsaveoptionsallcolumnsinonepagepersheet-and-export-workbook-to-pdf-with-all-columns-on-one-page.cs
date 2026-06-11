using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    public class ExportAllColumnsOnePage
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the worksheet with enough columns to demonstrate the option
                for (int col = 0; col < 50; col++)
                {
                    // Header row
                    sheet.Cells[0, col].PutValue($"Column {col + 1}");
                    // Sample data row
                    sheet.Cells[1, col].PutValue($"Data {col + 1}");
                }

                // Configure PDF save options to fit all columns on a single page per sheet
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    AllColumnsInOnePagePerSheet = true,
                    OnePagePerSheet = true // ensures the whole sheet fits on one page
                };

                string outputPath = "AllColumnsOnePage.pdf";

                // Save the workbook as PDF using the configured options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook exported to PDF with all columns on one page per sheet: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during export: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportAllColumnsOnePage.Run();
        }
    }
}