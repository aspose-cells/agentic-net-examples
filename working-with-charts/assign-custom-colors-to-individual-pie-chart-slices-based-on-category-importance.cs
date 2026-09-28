// Title: How to assign custom colors to individual slices of a pie chart in Excel using Aspose.Cells for .NET (C#)
// AI Prompts: Create a pie chart from worksheet data and set the foreground color of each slice (High, Medium, Low) to red, orange, and green with Aspose.Cells in C#. | Disable automatic color variation for a pie chart series and apply specific Color values to chart points via the Points collection using Aspose.Cells.
// Common Searches: C# Aspose.Cells assign red orange green colors to pie chart slices | How to turn off color variation in a pie chart using Aspose.Cells .NET | Customizing individual pie slice colors based on category importance in Excel with Aspose.Cells | Aspose.Cells example for coloring pie chart points programmatically
// Tags: pie chart slice custom colors Aspose.Cells C# | disable automatic color variation Aspose.Cells chart series | set point foreground color Aspose.Cells chart | Excel pie chart color customization .NET | category based slice coloring Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// // This example creates a workbook, adds category data, inserts a pie chart, disables automatic color variation, and assigns red, orange, and green to the High, Medium, and Low slices before saving the file.
class CustomPieSliceColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: categories and their values
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");

        sheet.Cells["A2"].PutValue("High");
        sheet.Cells["A3"].PutValue("Medium");
        sheet.Cells["A4"].PutValue("Low");

        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Add a pie chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIdx];

        // Bind the chart to the data range
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Disable automatic color variation so we can set colors manually
        chart.NSeries[0].IsColorVaried = false;

        // Assign custom colors to each slice based on importance
        // Slice 0 (High)   -> Red
        chart.NSeries[0].Points[0].Area.ForegroundColor = Color.Red;
        // Slice 1 (Medium) -> Orange
        chart.NSeries[0].Points[1].Area.ForegroundColor = Color.Orange;
        // Slice 2 (Low)    -> Green
        chart.NSeries[0].Points[2].Area.ForegroundColor = Color.Green;

        // Save the workbook with the customized pie chart
        workbook.Save("CustomPieSliceColors.xlsx");
    }
}
