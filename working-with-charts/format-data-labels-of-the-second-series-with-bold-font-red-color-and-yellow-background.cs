// Title: Format second series data labels with bold font, red text, and yellow background in an Aspose.Cells column chart (C#)
// Description: This Aspose.Cells for .NET example creates a workbook, adds a column chart with two series, enables data labels for the second series, and applies bold styling, red font color, and a yellow label background before saving the file as an XLSX workbook.
// Keywords: Aspose.Cells | C# chart formatting | column chart data labels | second series styling | bold font data label | red text color chart | yellow label background | .NET Excel chart example | chart series formatting | Aspose.Cells data label customization
// Common Searches: Aspose.Cells format data labels second series C# | How to set bold red font for chart data labels in Aspose.Cells | Add yellow background to data labels in Aspose.Cells column chart | C# example for styling specific series data labels in Excel chart | Aspose.Cells change data label appearance for one series
// Developer Intent: Apply bold text, red font color, and a yellow background to the data labels of the second series in an Aspose.Cells column chart using C#.
// Use Cases: Highlight a competitor’s values in a sales comparison column chart. | Draw attention to key performance metrics in a financial report by emphasizing specific series labels. | Create visually distinct data labels for a dashboard that separates primary and secondary data series.
// AI Prompts: Show C# code that formats the data labels of the second series in an Aspose.Cells column chart with bold font, red text, and a yellow background. | Provide an Aspose.Cells for .NET example to enable and style data labels for a specific series, including font weight, text color, and label background. | Explain how to customize chart data label appearance for one series in an Excel workbook using Aspose.Cells and C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// This Aspose.Cells for .NET example creates a workbook, adds a column chart with two series, enables data labels for the second series, and applies bold styling, red font color, and a yellow label background before saving the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for two series
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Add a column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIdx];

        // Add the two series to the chart
        chart.NSeries.Add("B2:B4", true); // first series
        chart.NSeries.Add("C2:C4", true); // second series
        chart.NSeries.CategoryData = "A2:A4";

        // Access the second series (index 1)
        Series secondSeries = chart.NSeries[1];

        // Enable data labels for the second series
        secondSeries.DataLabels.ShowValue = true;

        // Format data labels: bold font, red color, yellow background
        secondSeries.DataLabels.Font.IsBold = true;
        secondSeries.DataLabels.Font.Color = Color.Red;
        secondSeries.DataLabels.Area.BackgroundColor = Color.Yellow;

        // Save the workbook
        workbook.Save("FormattedDataLabels.xlsx", SaveFormat.Xlsx);
    }
}
