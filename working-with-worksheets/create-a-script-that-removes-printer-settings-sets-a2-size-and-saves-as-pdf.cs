using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class RemovePrinterSettingsSetA2AndSavePdf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the page setup for the worksheet
                PageSetup pageSetup = worksheet.PageSetup;

                // Remove any printer settings
                pageSetup.PrinterSettings = null;

                // Set the paper size to A2 (420 x 594 mm)
                pageSetup.PaperSize = PaperSizeType.PaperA2;

                // Ensure workbook default paper size matches
                workbook.Settings.PaperSize = PaperSizeType.PaperA2;

                // Create PDF save options (default optimization is Standard)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF file
                string outputPath = "Result_A2_NoPrinterSettings.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}