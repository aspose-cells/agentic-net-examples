// Title: Create a column chart with a bottom‑center legend that doesn’t overlap data using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that builds a column chart from A2:B4, places the legend at the bottom center of the plot area, and saves the workbook. | Adjust an existing Aspose.Cells chart in C# to move its legend to the bottom middle and programmatically verify that the legend does not cover any data series. | Provide a complete C# example that adds sample data, creates a column chart, sets Legend.Position = LegendPositionType.Bottom, and confirms the layout remains clear.
// Common Searches: how to set bottom center legend in Aspose.Cells chart C# | Aspose.Cells column chart legend overlapping data series | C# Aspose.Cells place legend below chart without covering series | center legend in plot area using Aspose.Cells .NET | verify chart layout after moving legend in Aspose.Cells
// Tags: Aspose.Cells column chart legend positioning | C# set chart legend bottom center | Aspose.Cells prevent legend overlap | chart layout customization Aspose.Cells | Aspose.Cells plot area legend alignment

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes sample category and value data to cells A1:B4, adds a column chart, binds the series to the data range, sets the legend position to Bottom (centered relative to the plot area) to avoid covering the chart, and saves the file as ChartWithCenteredLegend.xlsx.
class ChartLegendExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart.
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Position the legend at the bottom (centered relative to the plot area).
            chart.Legend.Position = LegendPositionType.Bottom;

            // Save the workbook to a file.
            string outputPath = "ChartWithCenteredLegend.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
