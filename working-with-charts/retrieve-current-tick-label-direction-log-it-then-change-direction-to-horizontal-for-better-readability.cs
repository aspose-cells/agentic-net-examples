// Title: Read and Set Chart Category Axis Tick Label Direction with Aspose.Cells for .NET
// Description: Shows how to retrieve the current TickLabels.DirectionType of a chart’s Category axis, log it to the console, change the orientation to Horizontal for better readability, and save the workbook.
// Keywords: Aspose.Cells | C# | Chart tick label direction | ChartTextDirectionType | CategoryAxis | Horizontal tick labels | .NET Excel chart | read tick label orientation | set tick label direction
// Common Searches: Aspose.Cells get tick label direction | change chart axis label orientation Aspose.Cells | ChartTextDirectionType Horizontal example | log chart tick label direction C# | modify category axis tick labels Aspose.Cells
// Developer Intent: Retrieve the current tick label orientation of a chart axis, output it, and set it to Horizontal.
// Use Cases: Improve readability of automatically generated Excel charts by forcing horizontal tick labels. | Capture the original label orientation for debugging or audit logs. | Dynamically adjust label direction based on the length of category names. | Standardize chart appearance across automated reporting pipelines.
// AI Prompts: Generate C# code using Aspose.Cells that reads the CategoryAxis.TickLabels.DirectionType, prints the value, sets it to Horizontal, and saves the workbook. | Provide an example that logs the existing tick label direction before modifying it in an Aspose.Cells chart. | Create a function that evaluates category label length and switches TickLabels.DirectionType between Rotate45 and Horizontal accordingly.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to retrieve the current TickLabels.DirectionType of a chart’s Category axis, log it to the console, change the orientation to Horizontal for better readability, and save the workbook.
class Program
{
    static void Main()
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
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Access the tick labels of the category (X) axis
        TickLabels tickLabels = chart.CategoryAxis.TickLabels;

        // Retrieve and log the current direction of the tick labels
        ChartTextDirectionType currentDirection = tickLabels.DirectionType;
        Console.WriteLine($"Current TickLabels DirectionType: {currentDirection}");

        // Change the direction to Horizontal for better readability
        tickLabels.DirectionType = ChartTextDirectionType.Horizontal;

        // Save the workbook to a file
        workbook.Save("TickLabelsDirectionDemo.xlsx");
    }
}
