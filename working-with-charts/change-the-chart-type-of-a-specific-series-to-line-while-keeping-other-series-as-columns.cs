using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ChangeSeriesChartType
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(210);

                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(90);
                sheet.Cells["C3"].PutValue(130);
                sheet.Cells["C4"].PutValue(160);
                sheet.Cells["C5"].PutValue(190);

                // Add a column chart (default type for all series)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add two series to the chart
                chart.NSeries.Add("B2:B5", true); // Series 1
                chart.NSeries.Add("C2:C5", true); // Series 2

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A5";

                // Change the first series to a line chart while keeping the second as column
                chart.NSeries[0].Type = ChartType.Line;

                // Save the workbook
                string outputPath = "ChangeSeriesChartType.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
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
            ChangeSeriesChartType.Run();
        }
    }
}