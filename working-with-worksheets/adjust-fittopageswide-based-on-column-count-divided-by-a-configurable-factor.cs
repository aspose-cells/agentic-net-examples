using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class AdjustFitToPagesWideDemo
    {
        // Configurable factor to divide the column count
        private const int Factor = 5; // Example: 5 columns per printed page

        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Sample data: populate some columns to demonstrate the logic
                // ------------------------------------------------------------
                for (int col = 0; col < 23; col++) // 23 columns of data
                {
                    worksheet.Cells[0, col].PutValue($"Header {col + 1}");
                    for (int row = 1; row <= 10; row++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
                    }
                }

                // ------------------------------------------------------------
                // Determine the number of used columns
                // MaxColumn is zero‑based, so add 1 to get the count
                // ------------------------------------------------------------
                int usedColumnCount = worksheet.Cells.MaxColumn + 1;

                // ------------------------------------------------------------
                // Calculate FitToPagesWide based on the configurable factor
                // Use Math.Max to ensure at least one page wide
                // ------------------------------------------------------------
                int fitToPagesWide = Math.Max(1, usedColumnCount / Factor);
                // If there is a remainder, round up to avoid truncation
                if (usedColumnCount % Factor != 0)
                    fitToPagesWide++;

                // Apply the page setup settings
                PageSetup pageSetup = worksheet.PageSetup;
                pageSetup.FitToPagesWide = fitToPagesWide; // Number of pages wide
                pageSetup.FitToPagesTall = 0;              // Let height adjust automatically

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                string outputPath = "AdjustedFitToPagesWide.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AdjustFitToPagesWideDemo.Run();
        }
    }
}