// Title: How to read a chart title's position ratios and convert them to 1/4000 units using Aspose.Cells for .NET
// Common Searches: Aspose.Cells get chart title XRatioToChart and YRatioToChart values in C# | convert chart title position ratio to 1/4000 units Aspose.Cells example | log chart title coordinates after creating a column chart with Aspose.Cells | retrieve chart title position in an Excel workbook using Aspose.Cells .NET | calculate chart title location in 1/4000 of chart area with Aspose.Cells
// Tags: chart title position ratios Aspose.Cells | title ratio to 4000 unit conversion | Aspose.Cells column chart title visibility | chart title coordinate extraction C# | Excel chart element position measurement

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart with a visible title, reads the title's XRatioToChart and YRatioToChart properties, converts each ratio to 1/4000‑unit values, writes both the raw ratios and converted positions to the console, and saves the file as ChartWithTitlePosition.xlsx.
class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a workbook and add sample data for the chart.
        // -----------------------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue("Jan");
        cells["A3"].PutValue("Feb");
        cells["A4"].PutValue("Mar");
        cells["B2"].PutValue(120);
        cells["B3"].PutValue(150);
        cells["B4"].PutValue(180);

        // -----------------------------------------------------------------
        // 2. Add a chart and configure its title.
        // -----------------------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Ensure the chart has a title and set some text.
        chart.Title.Text = "Quarterly Sales";
        // Optionally set the title to be visible (default is true).
        chart.Title.IsVisible = true;

        // -----------------------------------------------------------------
        // 3. Retrieve the title position ratios (relative to the chart area).
        //    These are expressed as fractions (0.0 – 1.0).
        // -----------------------------------------------------------------
        double xRatio = chart.Title.XRatioToChart; // Horizontal position ratio
        double yRatio = chart.Title.YRatioToChart; // Vertical position ratio

        // -----------------------------------------------------------------
        // 4. Convert the ratios to 1/4000 units.
        //    1 unit = 1/4000 of the chart area, therefore:
        //    positionInUnits = ratio * 4000
        // -----------------------------------------------------------------
        int xPositionUnits = (int)Math.Round(xRatio * 4000);
        int yPositionUnits = (int)Math.Round(yRatio * 4000);

        // -----------------------------------------------------------------
        // 5. Log the retrieved values.
        // -----------------------------------------------------------------
        Console.WriteLine("Chart Title Position (Ratios):");
        Console.WriteLine($"  X Ratio = {xRatio:F6}");
        Console.WriteLine($"  Y Ratio = {yRatio:F6}");
        Console.WriteLine();
        Console.WriteLine("Chart Title Position (1/4000 units):");
        Console.WriteLine($"  X Position = {xPositionUnits} (units of 1/4000 of chart area)");
        Console.WriteLine($"  Y Position = {yPositionUnits} (units of 1/4000 of chart area)");

        // -----------------------------------------------------------------
        // 6. Save the workbook (optional, just to verify the chart is correct).
        // -----------------------------------------------------------------
        workbook.Save("ChartWithTitlePosition.xlsx");
    }
}
