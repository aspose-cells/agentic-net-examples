// Title: Set Fixed Y‑Axis Range (0‑1000) and Major Unit for a Bar Chart with Aspose.Cells for .NET
// Description: Creates a workbook, adds a bar chart with sample data, disables automatic scaling, sets the Y‑axis minimum to 0, maximum to 1000, defines a major unit of 200, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells Y axis range | C# bar chart custom axis | disable automatic axis scaling | set chart min max values | major unit Aspose.Cells | fixed Y axis Aspose.Cells .NET | chart axis limits programmatically
// Common Searches: Aspose.Cells set Y axis minimum and maximum | C# bar chart Y axis 0 to 1000 Aspose.Cells | how to disable automatic axis scaling Aspose.Cells | set major unit for chart axis Aspose.Cells .NET | fixed Y axis range for Excel chart using Aspose
// Developer Intent: Programmatically define a fixed minimum, maximum, and major unit for the Y‑axis of a bar chart in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Standardize sales dashboards so every bar chart displays values from 0 to 1000, ensuring consistent visual comparison. | Generate regulatory reports that require a predefined Y‑axis range for compliance with formatting guidelines. | Create performance scorecards where multiple charts share the same axis limits to avoid misleading scale variations.
// AI Prompts: Write C# code with Aspose.Cells to set a bar chart Y‑axis minimum of 0, maximum of 1000, and major unit of 200. | Explain how to turn off automatic axis scaling and apply custom min/max values on a chart's value axis in Aspose.Cells for .NET. | Provide a step‑by‑step example that creates a bar chart, fixes the Y‑axis range, and saves the workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a bar chart with sample data, disables automatic scaling, sets the Y‑axis minimum to 0, maximum to 1000, defines a major unit of 200, and saves the file as an Excel workbook.
    public class AdjustYAxisScaleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the bar chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(200);
                worksheet.Cells["B3"].PutValue(600);
                worksheet.Cells["B4"].PutValue(950);

                // Add a bar chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Bar, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure the Y‑axis (value axis) to display from 0 to 1000
                Axis valueAxis = chart.ValueAxis;

                // Disable automatic min/max values
                valueAxis.IsAutomaticMinValue = false;
                valueAxis.IsAutomaticMaxValue = false;

                // Set explicit min and max values
                valueAxis.MinValue = 0;      // Minimum value
                valueAxis.MaxValue = 1000;   // Maximum value

                // Optional: set major unit for better tick spacing
                valueAxis.MajorUnit = 200;

                // Save the workbook with the configured chart
                string outputPath = "BarChart_With_CustomYAxis.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
            AdjustYAxisScaleDemo.Run();
        }
    }
}
