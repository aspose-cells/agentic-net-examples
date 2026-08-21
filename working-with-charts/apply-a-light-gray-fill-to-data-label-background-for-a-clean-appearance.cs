// Title: Set Light Gray Background for Chart Data Labels with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, enables data labels, applies a solid light‑gray fill to the label background, and saves the file as an .xlsx.
// Keywords: Aspose.Cells | C# | chart data label background | light gray fill | solid fill pattern | Excel chart formatting | set data label color
// Common Searches: Aspose.Cells set data label background color | how to apply solid fill to chart data labels in .NET | light gray data label background Aspose.Cells | change chart label background color C#
// Developer Intent: Set a solid light‑gray background for chart data labels using Aspose.Cells in C#.
// Use Cases: Generate column charts in automated reports where each data label has a light‑gray background for better readability. | Apply corporate styling to Excel workbooks by standardizing data label backgrounds across multiple charts. | Export dashboards with consistent label appearance, ensuring all chart data labels share the same light‑gray fill.
// AI Prompts: Show how to apply a solid light‑gray fill to chart data labels with Aspose.Cells for .NET. | Provide a C# example that customizes chart data label appearance, including background color and font, using Aspose.Cells. | Explain how to set different background colors for data labels of multiple series in an Aspose.Cells chart.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, inserts a column chart, enables data labels, applies a solid light‑gray fill to the label background, and saves the file as an .xlsx.
class ApplyLightGrayDataLabelBackground
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

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Apply a light gray fill to the data label background
        series.DataLabels.Area.FillFormat.Pattern = FillPattern.Solid;
        series.DataLabels.Area.BackgroundColor = Color.LightGray;

        // Save the workbook
        workbook.Save("ChartWithLightGrayDataLabels.xlsx");
    }
}
