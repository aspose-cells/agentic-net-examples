// Title: Detect X‑Axis Value Axis and Convert Column Chart to Scatter in Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart, checks if its X‑axis is a value axis (only Scatter or Bubble), and if not, switches the chart to Scatter, assigns numeric XValues, updates axis titles, and saves the file.
// Keywords: Aspose.Cells X axis value axis | convert column chart to scatter | numeric X axis Aspose.Cells | chart type detection Aspose.Cells | C# Aspose.Cells chart conversion | set XValues Aspose.Cells
// Common Searches: Aspose.Cells check if chart X axis is value axis | change column chart to scatter chart Aspose.Cells .NET | set XValues for scatter chart using Aspose.Cells | programmatically change chart type Aspose.Cells | numeric X axis in Aspose.Cells chart
// Developer Intent: Identify whether a chart’s X‑axis is a value axis and, when it isn’t, programmatically convert the chart to a Scatter type so numeric X data can be used.
// Use Cases: Validate chart axis before exporting to ensure numeric X values are displayed correctly. | Automatically convert category‑based charts to value‑based charts when source data contains numeric X values. | Generate dynamic reports where X values are timestamps or measurements that require a value axis. | Update axis titles after conversion to maintain clear chart labeling.
// AI Prompts: Generate C# code with Aspose.Cells that detects if a chart’s X axis is a value axis and converts a column chart to a scatter chart with XValues set to a cell range. | Show how to switch a chart type to Scatter and assign numeric X data using Aspose.Cells for .NET. | Explain the steps to programmatically change a chart’s X axis from category to value axis in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart, checks if its X‑axis is a value axis (only Scatter or Bubble), and if not, switches the chart to Scatter, assigns numeric XValues, updates axis titles, and saves the file.
    public class DetectAndConvertXAxis
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (numeric X values and Y values)
                sheet.Cells["A1"].PutValue("X");
                sheet.Cells["B1"].PutValue("Y");
                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue(i - 1);          // X = 1,2,3,4,5
                    sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // Y = 10,20,30,40,50
                }

                // Add a column chart (its X axis is a category axis by default)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B6", true);          // Y values
                chart.NSeries.CategoryData = "A2:A6";      // X values as categories

                // ----- Detect if the X axis is a value axis -----
                // For scatter and bubble charts the X axis is a value axis.
                bool isXAxisValueAxis = chart.Type == ChartType.Scatter ||
                                        chart.Type == ChartType.Bubble;

                Console.WriteLine("Is X axis a value axis? " + isXAxisValueAxis);

                // ----- If not, convert the chart to a scatter chart (numeric X axis) -----
                if (!isXAxisValueAxis)
                {
                    // Change chart type to Scatter which uses a value axis for X
                    chart.Type = ChartType.Scatter;

                    // Set the X values for the series to the numeric range
                    // (CategoryData is ignored for scatter charts; use XValues instead)
                    chart.NSeries[0].XValues = "A2:A6";

                    // Optionally, adjust axis titles for clarity
                    chart.CategoryAxis.Title.Text = "Numeric X Axis";
                    chart.ValueAxis.Title.Text = "Y Axis";

                    Console.WriteLine("Chart converted to Scatter. X axis now supports numeric data.");
                }

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DetectAndConvertXAxis.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DetectAndConvertXAxis.Run();
        }
    }
}
