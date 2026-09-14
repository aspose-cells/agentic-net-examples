// Title: Set custom minimum and maximum Y‑axis values for a column chart using Aspose.Cells in C#
// AI Prompts: Write C# code that creates a workbook, adds a column chart, and turns off automatic Y‑axis range calculation using Aspose.Cells. | Show how to assign specific MinValue, MaxValue, and MajorUnit to the chart's ValueAxis in Aspose.Cells for .NET. | Provide a complete example that configures a fixed Y‑axis range and saves the workbook as an XLSX file.
// Common Searches: Aspose.Cells C# set Y axis min value for column chart | disable automatic Y axis scaling Aspose.Cells .NET example | how to specify custom major unit on chart value axis using Aspose.Cells | fixed Y axis range for Excel chart generated with Aspose.Cells | set minimum and maximum values for chart axis programmatically in C#
// Tags: value axis manual limits Aspose.Cells | chart major unit configuration .NET | column chart Y axis customization C# | Aspose.Cells axis scaling control | Excel chart axis settings Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding a column chart with sample data, disabling automatic Y‑axis scaling, setting explicit MinValue, MaxValue, and MajorUnit on the ValueAxis, and saving the result as YAxisMinMaxDemo.xlsx.
    public class SetYAxisMinMaxDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
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

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the Y‑axis (value axis)
                Axis valueAxis = chart.ValueAxis;

                // Disable automatic min/max calculation
                valueAxis.IsAutomaticMinValue = false;
                valueAxis.IsAutomaticMaxValue = false;

                // Define custom minimum and maximum values
                valueAxis.MinValue = 5;   // Minimum Y‑axis value
                valueAxis.MaxValue = 60;  // Maximum Y‑axis value

                // Optionally set a major unit for clearer tick spacing
                valueAxis.IsAutomaticMajorUnit = false;
                valueAxis.MajorUnit = 10;

                // Save the workbook to a file
                string outputPath = "YAxisMinMaxDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
