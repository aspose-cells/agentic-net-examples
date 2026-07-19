// Title: Get and Convert Aspose.Cells Chart Title Position to 1/4000 Units (C#)
// Description: Creates a workbook, adds sample data and a column chart, sets a title, reads the title's XRatioToChart and YRatioToChart, converts each ratio to 1/4000‑unit coordinates, logs the raw and converted values, and saves the file.
// Keywords: Aspose.Cells chart title position | XRatioToChart C# | YRatioToChart conversion | 1/4000 unit coordinates | retrieve chart title coordinates .NET | Aspose.Cells chart layout
// Common Searches: Aspose.Cells get chart title XRatioToChart | convert chart title ratio to 1/4000 units | how to read chart title position Aspose.Cells | chart title coordinates in Excel using Aspose.Cells | Aspose.Cells chart title placement values
// Developer Intent: Read the X and Y ratio values of a chart title, transform them into 1/4000‑unit measurements, and output the results.
// Use Cases: Align shapes or annotations precisely relative to a chart title. | Store exact title placement for reproducible report layouts. | Log title coordinates to verify consistent formatting across generated workbooks.
// AI Prompts: Show how to set XRatioToChart and YRatioToChart to reposition a chart title in Aspose.Cells for .NET. | Provide code that converts chart title 1/4000‑unit ratios to pixel coordinates based on worksheet dimensions. | Explain how to retrieve and log title positions for multiple charts in a single workbook using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Creates a workbook, adds sample data and a column chart, sets a title, reads the title's XRatioToChart and YRatioToChart, converts each ratio to 1/4000‑unit coordinates, logs the raw and converted values, and saves the file.
class RetrieveChartTitlePosition
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set the chart title
        chart.Title.Text = "Sample Chart";

        // Retrieve the title position ratios (fraction of chart area)
        double xRatio = chart.Title.XRatioToChart;
        double yRatio = chart.Title.YRatioToChart;

        // Convert ratios to units of 1/4000 of the chart area
        int xUnits = (int)Math.Round(xRatio * 4000);
        int yUnits = (int)Math.Round(yRatio * 4000);

        // Log the retrieved and converted values
        Console.WriteLine($"Title X Ratio: {xRatio}");
        Console.WriteLine($"Title Y Ratio: {yRatio}");
        Console.WriteLine($"Title X in 1/4000 units: {xUnits}");
        Console.WriteLine($"Title Y in 1/4000 units: {yUnits}");

        // Save the workbook
        workbook.Save("ChartTitlePosition.xlsx");
    }
}
