// Title: Set a chart title programmatically with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, define its series and categories, make the chart title visible, assign the text "Sales Overview", and save the file as an Excel workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells chart title C# | set chart title programmatically | Aspose.Cells column chart example | chart title visibility .NET | Aspose.Cells workbook chart title
// Common Searches: Aspose.Cells how to add a chart title in C# | set chart title visible Aspose.Cells | C# Aspose.Cells column chart title text | programmatically change Excel chart title with Aspose | Aspose.Cells chart title example
// Developer Intent: Add or modify the title of a chart after it has been created using Aspose.Cells in C#.
// Use Cases: Label a sales column chart with a clear title before exporting to Excel. | Apply distinct titles to multiple charts in a single workbook to differentiate data sets. | Show or hide chart titles based on user settings when generating automated reports.
// AI Prompts: Generate C# code with Aspose.Cells to create a pie chart and set its title to "Market Share". | Show how to update an existing chart's title dynamically from a worksheet cell using Aspose.Cells. | Provide robust error handling for setting a chart title when the chart object may be null.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add sample data, insert a column chart, define its series and categories, make the chart title visible, assign the text "Sales Overview", and save the file as an Excel workbook using Aspose.Cells for C#.
class SetChartTitleDemo
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set the chart title programmatically
        chart.Title.IsVisible = true;          // Ensure the title is visible
        chart.Title.Text = "Sales Overview";   // Set the desired title text

        // Save the workbook with the chart
        workbook.Save("ChartWithTitle.xlsx");
    }
}
