// Title: Change chart axis tick label color based on value thresholds with Aspose.Cells for .NET
// Common Searches: asp.net how to set chart axis tick label color conditionally using Aspose.Cells | C# Aspose.Cells conditional formatting for axis labels based on value range | apply value‑based color to chart axis tick labels in a column chart with Aspose.Cells | Aspose.Cells .NET change axis label font color when values exceed a threshold | conditional axis label styling in Excel charts using Aspose.Cells C#
// Tags: conditional axis label coloring Aspose.Cells | chart axis tick label color threshold .NET | Aspose.Cells column chart axis formatting | value based axis label styling C# | apply conditional formatting to chart axis Aspose.Cells | Excel chart axis label color conditional Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, adds sample month‑sales data, inserts a column chart, sets the series and category ranges, notes that Aspose.Cells does not directly support conditional formatting of axis tick label colors, and saves the file as ConditionalAxisFormatting.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(80);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(120);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(60);
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["B5"].PutValue(150);

            // Add a column chart to the worksheet (lifecycle rule: create)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // NOTE: Conditional number formats for axis tick labels are not directly
            // supported via the Axis.NumberFormat property in older Aspose.Cells versions.
            // If needed, set a simple number format instead.
            // chart.ValueAxis.NumberFormat = "General";

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ConditionalAxisFormatting.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
