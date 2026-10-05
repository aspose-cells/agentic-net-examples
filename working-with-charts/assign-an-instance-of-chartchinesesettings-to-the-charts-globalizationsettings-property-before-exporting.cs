// Title: How to assign ChartChineseSettings to a chart’s GlobalizationSettings and export as PNG using Aspose.Cells in C#
// AI Prompts: Instantiate a Chinese chart globalization settings object, configure its Font and NumericStyle for Chinese locale, assign it to the chart's GlobalizationSettings, then call chart.ToImage to produce a PNG file. | Adjust the sample code to apply Chinese culture settings to the chart before invoking ToImage, ensuring the exported image reflects Chinese formatting.
// Common Searches: Aspose.Cells C# set Chinese number format for chart axis before exporting image | How to use ChartChineseSettings with Aspose.Cells chart globalization settings | Export Aspose.Cells chart to PNG with Chinese locale in .NET | Assign ChartChineseSettings to chart.GlobalizationSettings in C# example
// Tags: Aspose.Cells chart Chinese localization | chart localization configuration C# | generate PNG image from Aspose.Cells chart | set Chinese font for chart axis Aspose.Cells | apply Chinese locale numeric style to chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, builds a column chart, and exports it as a PNG image, but it omits assigning a ChartChineseSettings instance to the chart's GlobalizationSettings. Adding this step configures Chinese number formats and fonts before the image is generated.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Export the chart as a PNG image (default format is PNG)
            chart.ToImage("ChartOutput.png");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
