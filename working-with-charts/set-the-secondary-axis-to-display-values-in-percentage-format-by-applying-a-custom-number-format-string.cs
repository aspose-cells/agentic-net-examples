// Title: C# – Set secondary axis to percentage format in an Aspose.Cells column chart
// Description: The example builds an Excel workbook, inserts category labels and two series (primary values and ratio values), creates a column chart, assigns the ratio series to the secondary value axis, and formats that axis with the custom pattern "0.00%" before exporting the file.
// Keywords: Aspose.Cells | C# | secondary axis | percentage number format | custom number format string | column chart | chart tick label formatting | Excel chart API | secondary value axis formatting | Aspose.Cells chart example
// Common Searches: Aspose.Cells secondary axis percentage format | C# chart secondary value axis custom format | How to display secondary axis as percent in Aspose.Cells | Set tick label number format for secondary axis .NET | Column chart secondary axis formatting Aspose.Cells
// Developer Intent: Format the secondary value axis of a chart as a percentage.
// Use Cases: Financial statements where profit margin is plotted on a secondary axis and shown as %. | Business dashboard displaying growth rates alongside absolute sales figures, with growth shown as percent on the secondary axis. | Scientific report visualizing concentration ratios on a secondary axis, requiring percentage tick labels.
// AI Prompts: Provide C# code using Aspose.Cells to plot a series on the secondary axis and format its tick labels as percentages with two decimal places. | Show an Aspose.Cells example that creates a column chart and applies a custom "0.00%" number format to the secondary value axis. | Explain the steps to set a custom number format string for secondary axis labels in an Aspose.Cells chart (C#).

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example builds an Excel workbook, inserts category labels and two series (primary values and ratio values), creates a column chart, assigns the ratio series to the secondary value axis, and formats that axis with the custom pattern "0.00%" before exporting the file.
class SetSecondaryAxisPercentageFormat
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Primary Series");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Secondary Series");
        sheet.Cells["C2"].PutValue(0.1);   // 10%
        sheet.Cells["C3"].PutValue(0.25);  // 25%
        sheet.Cells["C4"].PutValue(0.5);   // 50%

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set data ranges for the two series
        chart.NSeries.Add("B2:B4", true);          // Primary series
        chart.NSeries.Add("C2:C4", true);          // Secondary series
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Apply a custom number format string to the secondary value axis tick labels
        chart.SecondValueAxis.TickLabels.NumberFormat = "0.00%";

        // Save the workbook
        workbook.Save("SecondaryAxisPercentageFormat.xlsx");
    }
}
