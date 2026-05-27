using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ClearPrinterSettingsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Clear any existing printer settings from the worksheet
                worksheet.PageSetup.PrinterSettings = null;

                // Apply new print configuration
                worksheet.PageSetup.PrintDraft = true;               // print without graphics
                worksheet.PageSetup.PrintQuality = 600;              // set print quality (dpi)
                worksheet.PageSetup.Orientation = PageOrientationType.Landscape;
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 1;

                // Add sample data
                worksheet.Cells["A1"].PutValue("Demo after clearing printer settings");
                worksheet.Cells["A2"].PutValue(DateTime.Now);

                // Define output file path
                string outputPath = "ClearedPrinterSettingsDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ClearPrinterSettingsDemo.Run();
        }
    }
}