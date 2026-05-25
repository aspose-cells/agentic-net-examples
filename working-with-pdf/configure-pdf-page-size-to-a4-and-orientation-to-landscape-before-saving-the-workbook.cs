using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfPageSetup
{
    public class ConfigurePdfPage
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Set default paper size to A4
                workbook.Settings.PaperSize = PaperSizeType.PaperA4;

                // Set orientation to Landscape
                workbook.Settings.SetPageOrientationType(PageOrientationType.Landscape);

                // Define output file path
                string outputPath = "output.pdf";

                // Save the workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ConfigurePdfPage.Run();
        }
    }
}