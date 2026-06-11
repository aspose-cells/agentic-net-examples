using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetPageOrderOverThenDown
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set page order to OverThenDown for proper pagination
                sheet.PageSetup.Order = PrintOrderType.OverThenDown;

                // Ensure portrait orientation (optional, default is portrait)
                sheet.PageSetup.Orientation = PageOrientationType.Portrait;

                // Define output file path
                string outputPath = "PageOrder_OverThenDown.xlsx";

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            SetPageOrderOverThenDown.Run();
        }
    }
}