// Title: Aspose.Cells .NET – Retrieve and Set Chart Category Axis Tick Labels Direction to Horizontal
// Description: Creates a workbook, adds a column chart, reads the current TickLabels.DirectionType of the category axis, logs it, changes the direction to Horizontal, and saves the file.
// Keywords: Aspose.Cells tick label direction | ChartTextDirectionType .NET | set category axis labels horizontal | read chart tick label orientation | Aspose.Cells chart axis formatting
// Common Searches: how to get tick label direction Aspose.Cells | change chart axis labels to horizontal C# | Aspose.Cells ChartTextDirectionType example | read and modify chart tick label orientation .NET | Aspose.Cells category axis label direction
// Developer Intent: Read the current tick‑label direction of a chart’s category axis, output it for diagnostics, then force the labels to a horizontal orientation.
// Use Cases: Log original label orientation before applying a layout change. | Ensure consistent, readable axis labels in automated Excel reports. | Adjust label direction dynamically based on chart size or data density.
// AI Prompts: Generate C# code that obtains TickLabels.DirectionType from a chart axis, prints the value, sets it to Horizontal, and saves the workbook using Aspose.Cells. | Explain how ChartTextDirectionType affects chart rendering and demonstrate switching between orientations programmatically in Aspose.Cells for .NET. | Write a routine that iterates over all charts in a workbook, logs each category axis tick‑label direction, and updates them to Horizontal.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart, reads the current TickLabels.DirectionType of the category axis, logs it, changes the direction to Horizontal, and saves the file.
public class TickLabelDirectionDemo
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

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set chart data source
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access tick labels of the category axis
            TickLabels tickLabels = chart.CategoryAxis.TickLabels;

            // Retrieve and log the current direction type
            ChartTextDirectionType currentDirection = tickLabels.DirectionType;
            Console.WriteLine($"Current Tick Labels Direction: {currentDirection}");

            // Change direction to Horizontal for better readability
            tickLabels.DirectionType = ChartTextDirectionType.Horizontal;

            // Save the workbook
            string outputPath = "TickLabelDirectionDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        TickLabelDirectionDemo.Run();
    }
}
