using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTemplate
{
    public class CreateChartWithLegendAndDataLabels
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart (topRow, leftColumn, bottomRow, rightColumn)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (including categories and values)
                chart.SetChartDataRange("A1:B4", true);

                // Enable data labels to show the value of each point
                chart.NSeries[0].DataLabels.ShowValue = true;

                // Set legend position (Bottom) – this placement does not overlay the chart
                chart.Legend.Position = LegendPositionType.Bottom;

                // Adjust legend size using the recommended properties (avoid obsolete Width/Height)
                chart.Legend.WidthRatioToChart = 0.8;   // 80% of chart width
                chart.Legend.HeightRatioToChart = 0.1;  // 10% of chart height
                chart.Legend.Font.Size = 12;
                chart.Legend.Font.IsBold = true;

                // Define output file path
                string outputPath = "ChartWithLegendAndDataLabels.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required by the project
    public class Program
    {
        public static void Main(string[] args)
        {
            CreateChartWithLegendAndDataLabels.Run();
        }
    }
}