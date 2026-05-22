using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class AssignSeriesToSecondaryAxis
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");

                worksheet.Cells["B1"].PutValue("Primary Series");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                worksheet.Cells["C1"].PutValue("Secondary Series");
                worksheet.Cells["C2"].PutValue(100);
                worksheet.Cells["C3"].PutValue(200);
                worksheet.Cells["C4"].PutValue(300);

                // Add a line chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Add two series: first on primary axis, second on secondary axis
                chart.NSeries.Add("B2:B4", true); // Primary series
                chart.NSeries.Add("C2:C4", true); // Secondary series
                chart.NSeries.CategoryData = "A2:A4";

                // Assign the second series to the secondary vertical axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Optional: customize the secondary axis (e.g., title)
                chart.SecondValueAxis.Title.Text = "Secondary Axis";

                // Define output file path
                string outputPath = "AssignSeriesToSecondaryAxis.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AssignSeriesToSecondaryAxis.Run();
        }
    }
}