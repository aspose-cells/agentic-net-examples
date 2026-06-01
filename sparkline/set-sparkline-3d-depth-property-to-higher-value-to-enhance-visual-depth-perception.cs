using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class Sparkline3DDepthDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a 3‑D chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(90);
                sheet.Cells["C3"].PutValue(110);
                sheet.Cells["C4"].PutValue(130);

                // Add a 3‑D column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];

                // Set data range for the chart
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Increase the depth of the 3‑D chart
                chart.DepthPercent = 300; // 20‑2000 range

                // Optional 3‑D visual adjustments
                chart.Perspective = 40;   // 0‑100
                chart.Elevation = 25;     // degrees

                // Save the workbook
                string outputPath = "Sparkline3DDepthDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            Sparkline3DDepthDemo.Run();
        }
    }
}