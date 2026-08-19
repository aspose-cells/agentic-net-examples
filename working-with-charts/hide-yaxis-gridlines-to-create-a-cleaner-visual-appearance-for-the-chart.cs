// Title: Hide Y‑Axis Gridlines in a Column Chart Using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, and disables the Y‑axis major gridlines via chart.ValueAxis.MajorGridLines.IsVisible = false, then saves the file.
// Keywords: Aspose.Cells hide Y axis gridlines | C# chart gridline visibility | Aspose.Cells column chart formatting | disable value axis gridlines Aspose.Cells | Aspose.Cells .NET chart appearance
// Common Searches: hide y‑axis gridlines Aspose.Cells C# | remove major gridlines from chart Aspose.Cells | Aspose.Cells hide chart gridlines .NET | C# Aspose.Cells chart formatting tutorial
// Developer Intent: Remove the Y‑axis (value axis) major gridlines from a column chart to achieve a cleaner visual layout.
// Use Cases: Produce sales reports where column charts have a minimalist look without Y‑axis gridlines. | Build Excel dashboards with multiple charts that hide Y‑axis gridlines to reduce visual clutter. | Export presentation‑ready visualizations that match corporate style guidelines by disabling gridlines.
// AI Prompts: Show how to hide both major and minor Y‑axis gridlines in an Aspose.Cells chart using C#. | Provide code to toggle X‑axis and Y‑axis gridline visibility with a boolean flag in Aspose.Cells for .NET. | Explain how to customize axis line colors and gridline visibility in Aspose.Cells charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart, and disables the Y‑axis major gridlines via chart.ValueAxis.MajorGridLines.IsVisible = false, then saves the file.
    public class HideYAxisGridlinesDemo
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
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart (rows 5‑20, columns 0‑8)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Hide Y‑axis (value axis) major gridlines for a cleaner look
                chart.ValueAxis.MajorGridLines.IsVisible = false;

                // Save the workbook
                workbook.Save("HideYAxisGridlinesDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
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
            HideYAxisGridlinesDemo.Run();
        }
    }
}
