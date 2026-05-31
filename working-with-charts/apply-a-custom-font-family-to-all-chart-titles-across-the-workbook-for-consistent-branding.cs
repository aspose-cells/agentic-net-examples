using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ApplyCustomFontToAllChartTitles
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for demonstration
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIdx1 = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart1 = sheet.Charts[chartIdx1];
                chart1.NSeries.Add("B2:B4", true);
                chart1.NSeries.CategoryData = "A2:A4";
                chart1.Title.Text = "First Chart";

                // Add a pie chart
                int chartIdx2 = sheet.Charts.Add(ChartType.Pie, 22, 0, 35, 8);
                Chart chart2 = sheet.Charts[chartIdx2];
                chart2.NSeries.Add("B2:B4", true);
                chart2.Title.Text = "Second Chart";

                // Custom font family for all chart titles
                string customFontFamily = "Calibri";

                // Apply custom font to every chart title in the workbook
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (Chart ch in ws.Charts)
                    {
                        ch.Title.IsVisible = true;               // Ensure title is shown
                        ch.Title.Font.Name = customFontFamily;   // Set font family
                        ch.Title.Font.Size = 14;                 // Optional size
                        ch.Title.Font.Color = Color.DarkBlue;    // Optional color
                    }
                }

                // Save the workbook
                string outputPath = "Workbook_With_CustomChartTitleFont.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyCustomFontToAllChartTitles.Run();
        }
    }
}