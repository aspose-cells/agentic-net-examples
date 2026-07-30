// Title: C# – Retrieve and Convert Aspose.Cells Chart Title Position (XRatioToChart/YRatioToChart) to 1/4000 Units
// Description: This Aspose.Cells for .NET example creates a workbook, adds a column chart with a title, reads the title's XRatioToChart and YRatioToChart (fraction of chart width/height), converts the ratios to the legacy 1/4000 unit scale, logs the values, and saves the file.
// Keywords: Aspose.Cells | C# | chart title position | XRatioToChart | YRatioToChart | 1/4000 units conversion | Excel chart title coordinates | retrieve chart title ratios | Aspose.Cells chart example
// Common Searches: Aspose.Cells get chart title XRatioToChart | How to read chart title position in Aspose.Cells .NET | Convert chart title ratios to 1/4000 units | Aspose.Cells chart title coordinates example | C# retrieve chart title placement
// Developer Intent: Obtain the relative X and Y position of a chart title, translate those ratios into the legacy 1/4000 unit system, and output the results.
// Use Cases: Fine‑tune chart title layout in automatically generated Excel reports. | Maintain compatibility with older Excel automation scripts that rely on 1/4000 unit positioning. | Log or audit title placement for quality‑control or debugging purposes. | Programmatically adjust title coordinates based on custom design rules.
// AI Prompts: Give a C# snippet that reads a chart title’s XRatioToChart and YRatioToChart, converts them to 1/4000 units, and writes the results to a log file using Aspose.Cells. | Explain how to reposition a chart title by setting XRatioToChart and YRatioToChart from desired 1/4000 unit values in Aspose.Cells for .NET. | Show code that iterates through all charts in a workbook, extracts each title’s position ratios, converts them to 1/4000 units, and stores the values in a dictionary.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This Aspose.Cells for .NET example creates a workbook, adds a column chart with a title, reads the title's XRatioToChart and YRatioToChart (fraction of chart width/height), converts the ratios to the legacy 1/4000 unit scale, logs the values, and saves the file.
class RetrieveChartTitlePosition
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

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set chart title
        chart.Title.Text = "Sample Chart Title";

        // Retrieve the title position ratios
        double xRatio = chart.Title.XRatioToChart; // fraction of chart width (0‑1)
        double yRatio = chart.Title.YRatioToChart; // fraction of chart height (0‑1)

        // Convert ratios to 1/4000 units (as used by the obsolete X/Y properties)
        int xIn4000 = (int)Math.Round(xRatio * 4000);
        int yIn4000 = (int)Math.Round(yRatio * 4000);

        // Log the values
        Console.WriteLine($"Title X Ratio: {xRatio}");
        Console.WriteLine($"Title Y Ratio: {yRatio}");
        Console.WriteLine($"Title X in 1/4000 units: {xIn4000}");
        Console.WriteLine($"Title Y in 1/4000 units: {yIn4000}");

        // Save the workbook (optional, just to demonstrate lifecycle)
        workbook.Save("RetrieveChartTitlePosition.xlsx");
    }
}
