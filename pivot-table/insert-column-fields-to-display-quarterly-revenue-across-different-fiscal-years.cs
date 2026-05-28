using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    public class InsertQuarterColumnsDemo
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Revenue";

            // Add the first column header for fiscal year
            sheet.Cells[0, 0].PutValue("FiscalYear");

            // Insert four columns after the first column to hold quarterly data
            sheet.Cells.InsertColumns(1, 4);

            // Get localized quarter names
            PivotGlobalizationSettings globalization = new PivotGlobalizationSettings();
            string[] quarterNames = globalization.GetTextOf4Quaters();

            // Set quarter names as column headers (B1 to E1)
            for (int i = 0; i < quarterNames.Length; i++)
            {
                sheet.Cells[0, i + 1].PutValue(quarterNames[i]);
            }

            // Populate sample fiscal year data and quarterly revenue
            int startYear = 2021;
            int rows = 5; // number of fiscal years to display
            Random rnd = new Random();

            for (int r = 0; r < rows; r++)
            {
                // Fiscal year column
                sheet.Cells[r + 1, 0].PutValue(startYear + r);

                // Quarterly revenue columns (random values for demonstration)
                for (int q = 0; q < quarterNames.Length; q++)
                {
                    double revenue = rnd.Next(5000, 20000);
                    sheet.Cells[r + 1, q + 1].PutValue(revenue);
                }
            }

            // Define output file path
            string outputPath = "QuarterlyRevenue.xlsx";

            // Save the workbook (overwrite if exists)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
        }
    }
}