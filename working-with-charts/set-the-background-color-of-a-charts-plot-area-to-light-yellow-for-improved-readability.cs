// Title: Set Chart Plot Area Background to Light Yellow with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, and changes the chart's plot area fill to LightYellow using Aspose.Cells before saving as ChartPlotAreaBackground.xlsx.
// Keywords: Aspose.Cells | C# chart formatting | plot area background color | light yellow chart fill | Excel chart styling .NET
// Common Searches: Aspose.Cells change chart plot area color | C# set plot area background light yellow | how to fill chart plot area Aspose.Cells | Excel chart background color using Aspose
// Developer Intent: Apply a LightYellow fill to a chart's plot area in an Aspose.Cells workbook.
// Use Cases: Enhance readability of column charts with a subtle background shade. | Standardize chart appearance across generated Excel reports. | Programmatically style multiple charts in a workbook for branding consistency.
// AI Prompts: Generate C# code that sets the plot area background of any Aspose.Cells chart to a custom color. | Explain how to conditionally assign different background shades to chart plot areas using Aspose.Cells. | Show a loop that applies a LightYellow fill to the plot area of all charts in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// Creates a workbook, adds sample data, inserts a column chart, and changes the chart's plot area fill to LightYellow using Aspose.Cells before saving as ChartPlotAreaBackground.xlsx.
class SetPlotAreaBackground
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
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
        chart.SetChartDataRange("A1:B4", true);

        // Set the background color of the plot area to light yellow
        chart.PlotArea.Area.BackgroundColor = Color.LightYellow;

        // Save the workbook
        workbook.Save("ChartPlotAreaBackground.xlsx");
    }
}
