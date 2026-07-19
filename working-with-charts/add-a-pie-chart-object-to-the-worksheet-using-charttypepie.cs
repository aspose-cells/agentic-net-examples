// Title: Create a Pie Chart in Excel with Aspose.Cells for .NET (C#)
// Description: This C# example builds a new workbook, writes category and value data to cells A1:B4, inserts a ChartType.Pie chart (rows 5‑15, columns 0‑5), links the series to B2:B4 and categories to A2:A4, adds a title, and saves the file as PieChart.xlsx.
// Keywords: Aspose.Cells pie chart C# | ChartType.Pie example .NET | Excel pie chart automation | add chart to worksheet Aspose.Cells | C# generate Excel chart | Aspose.Cells chart data series | Excel workbook save with chart
// Common Searches: how to add a pie chart with Aspose.Cells for .NET | Aspose.Cells C# code for ChartType.Pie | set category and values for an Excel pie chart using Aspose | example of creating a pie chart in a workbook with Aspose.Cells
// Developer Intent: Programmatically insert and configure a pie chart in an Excel file using Aspose.Cells for .NET.
// Use Cases: Produce a sales‑by‑product pie chart for a monthly report. | Automate dashboard generation by adding titled pie charts to exported workbooks. | Create multiple pie charts across worksheets to compare regional performance.
// AI Prompts: Convert the pie chart in the example to a donut chart with Aspose.Cells for .NET. | Show how to customize slice colors, add a legend, and format data labels for the pie chart. | Explain how to bind the chart to a dynamic named range so it updates when worksheet data changes.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example builds a new workbook, writes category and value data to cells A1:B4, inserts a ChartType.Pie chart (rows 5‑15, columns 0‑5), links the series to B2:B4 and categories to A2:A4, adds a title, and saves the file as PieChart.xlsx.
class AddPieChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Add a pie chart to the worksheet (topRow, leftColumn, bottomRow, rightColumn)
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Optional: set a title for the chart
        chart.Title.Text = "Fruit Distribution";

        // Save the workbook with the chart
        workbook.Save("PieChart.xlsx", SaveFormat.Xlsx);
    }
}
