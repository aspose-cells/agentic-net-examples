using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ModifyChartTitle
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set an initial title
                chart.Title.Text = "Original Chart Title";
                chart.Title.IsVisible = true;

                // Read and display the current title
                Console.WriteLine("Current chart title: " + chart.Title.Text);

                // Modify the title to a custom string
                chart.Title.Text = "Custom Chart Title";

                // Verify the change
                Console.WriteLine("Updated chart title: " + chart.Title.Text);

                // Save the workbook
                string outputPath = "ModifiedChartTitle.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ModifyChartTitle.Run();
        }
    }
}