using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class YAxisScalingDemo
    {
        public static void Run()
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
            worksheet.Cells["B4"].PutValue(55);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Obtain the Y‑axis (value axis) of the chart
            Axis valueAxis = chart.ValueAxis;

            // Disable automatic min/max calculation
            valueAxis.IsAutomaticMinValue = false;
            valueAxis.IsAutomaticMaxValue = false;

            // Define custom minimum and maximum values for the Y‑axis
            valueAxis.MinValue = 5;   // Minimum value displayed on the axis
            valueAxis.MaxValue = 60;  // Maximum value displayed on the axis

            // Optionally set major unit for clearer tick marks
            valueAxis.IsAutomaticMajorUnit = false;
            valueAxis.MajorUnit = 10;

            // Save the workbook to a file
            string outputPath = "YAxisScalingDemo.xlsx";
            workbook.Save(outputPath);
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main()
        {
            try
            {
                YAxisScalingDemo.Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}