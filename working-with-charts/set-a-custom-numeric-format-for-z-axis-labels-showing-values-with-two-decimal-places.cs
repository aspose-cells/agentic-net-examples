using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ZAxisTickLabelsNumberFormatDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(1.2345);
                worksheet.Cells["B3"].PutValue(2.3456);
                worksheet.Cells["B4"].PutValue(3.4567);

                // Add a 3‑D column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";     // Categories

                // Set custom numeric format for the value (Y) axis tick labels (two decimal places)
                // ZAxis is not available in older Aspose.Cells versions, so we format the Y axis instead.
                if (chart.ValueAxis?.TickLabels != null)
                {
                    chart.ValueAxis.TickLabels.NumberFormat = "0.00";
                }

                // Determine output path and ensure the directory exists
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ZAxisTickLabelsNumberFormatDemo.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            ZAxisTickLabelsNumberFormatDemo.Run();
        }
    }
}