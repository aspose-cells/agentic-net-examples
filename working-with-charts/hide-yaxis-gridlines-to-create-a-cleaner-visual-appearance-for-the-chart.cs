// Title: C# example for hiding major and minor Y‑axis gridlines in an Aspose.Cells column chart
// Common Searches: C# hide Y axis gridlines in Aspose.Cells column chart | Aspose.Cells hide value axis major gridlines .NET | remove minor gridlines from chart axis using Aspose.Cells C# | how to disable chart gridlines in an Excel workbook with Aspose.Cells | Aspose.Cells chart formatting hide Y‑axis gridlines example
// Tags: aspnet hide chart value axis gridlines | aspnet column chart gridline visibility | aspnet chart formatting hide major gridlines | aspnet chart formatting hide minor gridlines | excel workbook chart appearance customization

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a new workbook, adds sample sales data, inserts a column chart, and then disables both major and minor gridlines on the chart's Y‑axis (value axis) before saving the file as ChartHideYGridlines.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);

            // Add a column chart to the worksheet (rows 5‑20, columns 0‑7)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide Y‑axis (value axis) major and minor gridlines for a cleaner look
            chart.ValueAxis.MajorGridLines.IsVisible = false;
            chart.ValueAxis.MinorGridLines.IsVisible = false;

            // Save the workbook with the chart
            workbook.Save("ChartHideYGridlines.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
