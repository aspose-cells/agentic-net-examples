// Title: C# – Set Fixed Y‑Axis (0‑100) for a Column Chart with Aspose.Cells
// Description: This example creates a new workbook, adds sample data, inserts a column chart, turns off automatic scaling on the value axis, assigns a minimum of 0 and a maximum of 100, and saves the file as an Excel workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart axis range | C# set chart Y axis min max | disable automatic axis scaling Aspose.Cells | fixed 0‑100 axis column chart .NET | standardized chart scaling Aspose | Excel chart axis customization C#
// Common Searches: how to fix chart y‑axis limits in Aspose.Cells | Aspose.Cells set column chart axis to 0‑100 | disable auto scaling for chart axis .NET | standardize axis range across multiple Excel charts
// Developer Intent: Apply a custom scale of 0 to 100 on the value (Y) axis of a column chart created with Aspose.Cells.
// Use Cases: Build dashboards where every chart shares a uniform 0‑100 scale for easy visual comparison. | Generate financial or KPI reports that require a consistent axis regardless of data fluctuations. | Create templates for automated Excel exports that must maintain the same Y‑axis range across all generated charts.
// AI Prompts: Generate C# code with Aspose.Cells that creates a line chart and forces the Y‑axis to start at 0 and end at 100. | Show how to disable automatic scaling for both primary and secondary axes in a stacked column chart using Aspose.Cells. | Explain the steps to retrieve a chart’s axis object after adding the chart and set custom min/max values in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, adds sample data, inserts a column chart, turns off automatic scaling on the value axis, assigns a minimum of 0 and a maximum of 100, and saves the file as an Excel workbook using Aspose.Cells for .NET.
    class SetAxisScaling
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
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(50);
                sheet.Cells["B4"].PutValue(90);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Disable automatic scaling and set standardized min/max values
                Axis valueAxis = chart.ValueAxis;
                valueAxis.IsAutomaticMinValue = false;
                valueAxis.IsAutomaticMaxValue = false;
                valueAxis.MinValue = 0;    // Minimum value set to 0
                valueAxis.MaxValue = 100;  // Maximum value set to 100

                // Save the workbook
                workbook.Save("ChartWithStandardizedScaling.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SetAxisScaling.Run();
        }
    }
}
