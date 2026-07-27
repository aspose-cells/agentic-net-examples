// Title: Read and Set Chart Tick Label Direction to Horizontal with AspNet Aspose.Cells
// Description: C# example that creates a workbook, adds a column chart, reads the current ChartTextDirectionType of the category axis tick labels, logs it to the console, switches the direction to Horizontal for better readability, and saves the file.
// Keywords: Aspose.Cells tick label direction | ChartTextDirectionType Horizontal | C# Aspose.Cells chart axis label orientation | read chart tick labels Aspose.Cells | set category axis label direction .NET | Excel chart label readability | Aspose.Cells chart formatting
// Common Searches: how to get tick label direction Aspose.Cells | set chart axis labels horizontal Aspose.Cells C# | ChartTextDirectionType example .NET | log current tick label orientation Aspose.Cells | change X axis label direction programmatically
// Developer Intent: Retrieve the existing tick‑label orientation of a chart’s category axis, output it, then force the labels to a horizontal layout.
// Use Cases: Standardize X‑axis label orientation in automated Excel reports. | Audit existing chart label settings before applying a new style. | Dynamically adjust label direction based on category name length or locale.
// AI Prompts: Generate C# code with Aspose.Cells that reads the current tick label direction of a chart’s category axis and prints it. | Show how to change the TickLabels.DirectionType to Horizontal and save the workbook using Aspose.Cells for .NET. | Explain a conditional approach to set tick label direction based on the length of category labels in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTickLabelDirectionDemo
{
    // C# example that creates a workbook, adds a column chart, reads the current ChartTextDirectionType of the category axis tick labels, logs it to the console, switches the direction to Horizontal for better readability, and saves the file.
    class Program
    {
        static void Main()
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

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the tick labels of the category (X) axis
            TickLabels tickLabels = chart.CategoryAxis.TickLabels;

            // Retrieve and log the current direction type
            ChartTextDirectionType currentDirection = tickLabels.DirectionType;
            Console.WriteLine($"Current Tick Labels Direction: {currentDirection}");

            // Change the direction to Horizontal for better readability
            tickLabels.DirectionType = ChartTextDirectionType.Horizontal;

            // Save the workbook
            workbook.Save("TickLabelsDirectionDemo.xlsx");
        }
    }
}
