// Title: Detect X‑Axis Value Axis and Convert a Column Chart to Scatter with Numeric X in Aspose.Cells for .NET
// Description: Creates a workbook with a column chart, checks whether the X‑axis is a value axis (Scatter/Bubble), and if it isn’t, switches the chart to a Scatter type, assigns a numeric X‑value range, and saves the file.
// Keywords: Aspose.Cells X axis value axis detection | convert column chart to scatter Aspose.Cells | set numeric X values Aspose.Cells .NET | chart axis conversion Aspose.Cells | programmatic chart type change Aspose
// Common Searches: how to check if chart X axis is a value axis in Aspose.Cells | convert category axis chart to scatter chart .NET | assign custom X values to series Aspose.Cells | change chart type programmatically Aspose.Cells | numeric X axis for Excel chart using Aspose
// Developer Intent: Determine if a chart’s X‑axis is a value axis and, when it isn’t, change the chart to a Scatter type and bind numeric X‑values.
// Use Cases: Validate the axis type of an automatically generated column chart before exporting. | Transform a categorical X‑axis into a numeric axis for statistical analysis. | Generate Excel reports that start with category data and later require numeric positioning without manual user intervention.
// AI Prompts: Generate C# code with Aspose.Cells that detects a chart’s X‑axis type and converts it to a Scatter chart with custom numeric X values. | Show how to programmatically change a column chart to a scatter chart and set the XValues property in Aspose.Cells for .NET. | Explain the steps to identify a value‑axis chart and modify its axis configuration to support numeric data using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAxisConversion
{
    // Creates a workbook with a column chart, checks whether the X‑axis is a value axis (Scatter/Bubble), and if it isn’t, switches the chart to a Scatter type, assigns a numeric X‑value range, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (numeric Y values, categorical X values)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart (X axis is a Category axis by default)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                // Category data is automatically taken from the adjacent column (A2:A4)

                // ------------------------------------------------------------
                // Detect if the X axis is a Value axis.
                // For most chart types the X axis is a Category axis.
                // Scatter and Bubble charts use a Value axis for X.
                // ------------------------------------------------------------
                bool isXAxisValueAxis = chart.Type == ChartType.Scatter ||
                                        chart.Type == ChartType.Bubble;

                Console.WriteLine("Initial chart type: " + chart.Type);
                Console.WriteLine("Is X axis a Value axis? " + isXAxisValueAxis);

                // ------------------------------------------------------------
                // If the X axis is not a Value axis, convert the chart to a
                // Scatter chart which uses a numeric X axis, and assign numeric
                // X values to the series.
                // ------------------------------------------------------------
                if (!isXAxisValueAxis)
                {
                    // Change chart type to Scatter (numeric X axis)
                    chart.Type = ChartType.Scatter;

                    // Provide numeric X values (e.g., 1, 2, 3) for the series
                    sheet.Cells["C2"].PutValue(1);
                    sheet.Cells["C3"].PutValue(2);
                    sheet.Cells["C4"].PutValue(3);

                    // Assign the X values range to the series
                    chart.NSeries[0].XValues = "C2:C4";

                    Console.WriteLine("Chart converted to Scatter. X axis now supports numeric data.");
                }

                // Save the workbook
                string outputPath = "ChartAxisConversion.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved as " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
