using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class SetYAxisMinMaxDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(30);
                worksheet.Cells["B4"].PutValue(50);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Obtain the value (Y) axis of the chart
                Axis valueAxis = chart.ValueAxis;

                // Turn off automatic min/max calculation
                valueAxis.IsAutomaticMinValue = false;
                valueAxis.IsAutomaticMaxValue = false;

                // Define custom minimum and maximum values for the Y‑axis
                valueAxis.MinValue = 5;   // Minimum Y value
                valueAxis.MaxValue = 60;  // Maximum Y value

                // Optionally set a custom major unit for tick spacing
                valueAxis.IsAutomaticMajorUnit = false;
                valueAxis.MajorUnit = 10;

                // Define output file path
                string outputPath = "YAxisMinMaxDemo.xlsx";

                // Save the workbook with the configured chart
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetYAxisMinMaxDemo.Run();
        }
    }
}