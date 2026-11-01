// Title: Hide Y‑axis major gridlines in a column chart with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a column chart and turn off the Y‑axis major gridlines. | Show how to set the ValueAxis.MajorGridLines.IsVisible property to false for a chart in Aspose.Cells .NET. | Provide a complete Aspose.Cells example that builds a workbook, adds data, inserts a column chart, and hides the Y‑axis gridlines before saving.
// Common Searches: Aspose.Cells C# hide Y axis gridlines in column chart | disable value axis major gridlines Aspose.Cells .NET example | remove Y‑axis lines from Excel chart using Aspose.Cells C# | chart formatting hide Y‑axis gridlines Aspose.Cells tutorial | how to turn off Y axis gridlines in Aspose.Cells generated Excel file
// Tags: hide value axis major gridlines Aspose.Cells C# | column chart gridline visibility Aspose.Cells | Aspose.Cells chart formatting hide Y axis | set ValueAxis.MajorGridLines.IsVisible false | Aspose.Cells workbook chart appearance customization

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a new workbook, fills it with sample sales data, adds a column chart, disables the Y‑axis major gridlines by setting ValueAxis.MajorGridLines.IsVisible to false, and saves the file as ChartWithHiddenYGridlines.xlsx.
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

            // Add a column chart (positioned from row 5, column 0 to row 20, column 7)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the series (values)
            chart.NSeries.Add("B2:B4", true);
            // Category (X‑axis) data – optional; if omitted, default numeric categories are used
            // chart.NSeries[0].CategoryData = "A2:A4";

            // Hide Y‑axis (value axis) major gridlines for a cleaner look
            chart.ValueAxis.MajorGridLines.IsVisible = false;

            // Save the workbook
            string outputPath = "ChartWithHiddenYGridlines.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
