// Title: C# Aspose.Cells – Safely Set Leader Lines on Charts with Exception Handling
// Description: Demonstrates how to add a column chart in Aspose.Cells, attempt to enable leader lines, and gracefully catch the CellsException thrown for unsupported chart types. The sample wraps leader‑line configuration and workbook saving in try‑catch blocks, logs errors, and shows how to continue processing without a crash.
// Keywords: Aspose.Cells | C# | leader lines | chart exception handling | unsupported chart type | CellsException | column chart | Excel chart API | error handling | Aspose.Cells chart | leader line properties | try catch
// Common Searches: Aspose.Cells catch exception for leader lines | which chart types support leader lines in Aspose.Cells | C# example handling unsupported chart features Aspose.Cells | error handling when setting leader lines on a column chart | Aspose.Cells leader lines not supported column chart
// Developer Intent: Add robust try‑catch logic around leader‑line settings to prevent runtime failures on chart types that do not support them.
// Use Cases: Prevent application crashes by catching CellsException when configuring leader lines on incompatible charts. | Log detailed error messages and optionally switch to a chart type that supports leader lines. | Ensure the workbook is saved even if chart configuration fails, using separate error handling for the save operation.
// AI Prompts: Generate C# code that checks a chart's type before enabling HasLeaderLines in Aspose.Cells. | Show how to catch Aspose.Cells.CellsException when setting leader line properties on a series. | Provide a fallback routine that replaces a column chart with a line chart if leader line configuration throws an exception.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing; // For drawing related enums and classes

namespace AsposeCellsExamples
{
    // Demonstrates how to add a column chart in Aspose.Cells, attempt to enable leader lines, and gracefully catch the CellsException thrown for unsupported chart types. The sample wraps leader‑line configuration and workbook saving in try‑catch blocks, logs errors, and shows how to continue processing without a crash.
    public class LeaderLinesErrorHandlingDemo
    {
        // Entry point required for compilation
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
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

            // Add a column chart (leader lines are not supported for column charts)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Attempt to configure leader lines – this will throw on unsupported chart types
            try
            {
                Series series = chart.NSeries[0];
                series.HasLeaderLines = true;               // Enable leader lines
                series.LeaderLines.IsAuto = false;          // Disable automatic formatting
                // The Style property may not be available in some versions; omit if not supported
                // series.LeaderLines.Style = Aspose.Cells.Drawing.LineStyle.Dot;
                series.LeaderLines.WeightPt = 1.5;          // Set line weight
                series.LeaderLines.Color = Color.Blue;      // Set line color

                Console.WriteLine("Leader lines configured successfully.");
            }
            catch (Exception ex)
            {
                // Handle the exception – typically Aspose.Cells.CellsException
                Console.WriteLine($"Error configuring leader lines: {ex.Message}");
                // Optionally, fallback to a supported chart type or skip configuration
            }

            // Save the workbook with safety check
            try
            {
                string outputPath = "LeaderLinesErrorHandlingDemo_out.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
