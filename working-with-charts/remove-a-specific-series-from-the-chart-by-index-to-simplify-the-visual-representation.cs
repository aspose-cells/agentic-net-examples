using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

namespace AsposeCellsExamples
{
    public class RemoveSeriesByIndexDemo
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
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(300);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(150);
                sheet.Cells["C3"].PutValue(250);
                sheet.Cells["C4"].PutValue(350);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Add two series to the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Remove the first series (index 0)
                chart.NSeries.RemoveAt(0);

                // Output the remaining series count for verification
                Console.WriteLine($"Remaining series count: {chart.NSeries.Count}");

                // Save the workbook
                string outputPath = "RemoveSeriesByIndexDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveSeriesByIndexDemo.Run();
        }
    }
}