using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDemo
{
    public class RefreshChartDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate initial data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart and set its data range
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Initial calculation (optional, ensures chart is rendered correctly the first time)
                chart.Calculate();

                // ----- Update source cells -----
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["B3"].PutValue(25);
                sheet.Cells["B4"].PutValue(35);

                // Refresh the chart so it reflects the updated cell values
                chart.Calculate();

                // Define output file path
                string outputPath = "RefreshedChart.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (chart cache will be refreshed automatically on save)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required for console applications
    public class Program
    {
        public static void Main(string[] args)
        {
            RefreshChartDemo.Run();
        }
    }
}