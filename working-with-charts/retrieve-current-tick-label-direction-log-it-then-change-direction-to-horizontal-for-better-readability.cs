// Title: Read current tick label direction and set it to horizontal on a chart’s category axis using Aspose.Cells for .NET (C#)
// AI Prompts: Demonstrate how to retrieve the DirectionType of a chart’s category axis tick labels and print it to the console with Aspose.Cells in C#. | Show C# code that changes the category axis tick labels to a horizontal orientation in an Aspose.Cells chart. | Provide a complete example that creates a workbook, adds a column chart, logs the existing tick label direction, updates it to horizontal, and saves the file.
// Common Searches: Aspose.Cells C# get category axis tick label orientation | How to change chart axis tick labels to horizontal in Aspose.Cells .NET | Log current tick label DirectionType before modifying chart with Aspose.Cells | Set chart tick label direction horizontal programmatically using Aspose.Cells | Read and update chart axis label text direction Aspose.Cells example
// Tags: Aspose.Cells chart category axis tick label direction | C# set chart tick labels horizontal | read chart tick label DirectionType .NET | modify chart axis label orientation Aspose.Cells | save workbook after changing tick label direction

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, reads the current DirectionType of the category axis tick labels, writes the value to the console, changes the direction to Horizontal for better readability, and saves the workbook as TickLabelsDirectionChanged.xlsx.
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

        // Save the workbook
        workbook.Save("TickLabelsDirectionChanged.xlsx");
    }
}
