using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ApplyMarginsAndFitToOnePage
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Set custom margins of 0.5 inches on all sides
            pageSetup.TopMarginInch = 0.5;
            pageSetup.BottomMarginInch = 0.5;
            pageSetup.LeftMarginInch = 0.5;
            pageSetup.RightMarginInch = 0.5;

            // Scale the worksheet to fit on one page (both width and height)
            pageSetup.SetFitToPages(1, 1);

            // Define output file path
            string outputPath = "ScaledWithMargins.xlsx";

            // Save the workbook (overwrites if exists)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}