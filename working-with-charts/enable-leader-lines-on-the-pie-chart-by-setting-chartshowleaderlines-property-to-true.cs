// Title: Enable Leader Lines on a Pie Chart Using Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, populate it with category/value data, insert a pie chart, bind the series, enable leader lines by setting HasLeaderLines = true, and save the workbook as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | pie chart | leader lines | HasLeaderLines | chart series formatting | Excel chart automation | programmatic chart styling
// Common Searches: Aspose.Cells enable leader lines pie chart | C# set leader lines for Excel chart | How to show leader lines on a pie chart with Aspose.Cells | Chart series HasLeaderLines property .NET | Add leader lines to pie chart programmatically
// Developer Intent: Turn on leader lines for a pie‑chart series in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate sales‑distribution pie charts where each slice label is linked with a clear leader line. | Automate financial dashboards that require readable pie‑chart labels without overlap. | Apply consistent leader‑line styling to multiple pie charts across worksheets in a reporting pipeline.
// AI Prompts: Write C# code with Aspose.Cells that adds a doughnut chart and enables leader lines for every series. | Provide a function that toggles leader lines on a pie chart based on a boolean argument. | Explain how to customize leader‑line color, thickness, and dash style for a pie chart series using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, populate it with category/value data, insert a pie chart, bind the series, enable leader lines by setting HasLeaderLines = true, and save the workbook as an Excel file.
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

        // Enable leader lines for the first series of the chart
        chart.NSeries[0].HasLeaderLines = true;

        // Save the workbook to a file
        workbook.Save("EnableLeaderLinesDemo.xlsx");
    }
}
