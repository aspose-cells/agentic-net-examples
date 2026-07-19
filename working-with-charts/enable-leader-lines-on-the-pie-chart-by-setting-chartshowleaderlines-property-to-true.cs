// Title: Enable Leader Lines on a Pie Chart with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills it with category/value data, adds a pie chart, and activates leader lines by setting the series HasLeaderLines property to true. The workbook is saved as an Excel file showing connected labels on the pie slices.
// Keywords: Aspose.Cells | C# pie chart | leader lines | HasLeaderLines property | chart series formatting | .NET Excel chart | Aspose.Cells example | pie chart label connectors
// Common Searches: Aspose.Cells enable leader lines pie chart | C# set HasLeaderLines for pie chart series | show leader lines on Excel pie chart using Aspose.Cells | Aspose.Cells chart label connectors example | how to add leader lines to pie chart in .NET
// Developer Intent: Turn on leader lines for a pie‑chart series in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate reports where pie‑chart labels need clear connectors for readability. | Programmatically update existing workbooks to display leader lines on all pie charts. | Create automated Excel dashboards with pie slices linked to their category names via leader lines.
// AI Prompts: Write C# code that iterates over every series in a pie chart and sets HasLeaderLines = true with Aspose.Cells. | Show how to toggle leader lines on a pie chart based on a Boolean parameter and save the workbook. | Explain the visual effect of enabling HasLeaderLines on a pie chart and best practices for label placement in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills it with category/value data, adds a pie chart, and activates leader lines by setting the series HasLeaderLines property to true. The workbook is saved as an Excel file showing connected labels on the pie slices.
class EnableLeaderLinesDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pie chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a pie chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable leader lines for the first series
        chart.NSeries[0].HasLeaderLines = true;

        // Save the workbook to a file
        workbook.Save("PieChart_With_LeaderLines.xlsx");
    }
}
