using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class AdjustYAxisScaleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for a bar (column) chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(200);
                worksheet.Cells["B3"].PutValue(600);
                worksheet.Cells["B4"].PutValue(950);

                // Add a column chart (used as a bar chart) to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Disable automatic min/max values and set custom range 0 – 1000
                Axis valueAxis = chart.ValueAxis;
                valueAxis.IsAutomaticMinValue = false;
                valueAxis.MinValue = 0;          // Minimum value on Y‑axis
                valueAxis.IsAutomaticMaxValue = false;
                valueAxis.MaxValue = 1000;       // Maximum value on Y‑axis

                // Optional: set major unit for clearer grid lines
                valueAxis.MajorUnit = 200;

                // Save the workbook with the adjusted chart
                string outputPath = "AdjustedYAxisScale.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
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
            AdjustYAxisScaleDemo.Run();
        }
    }
}