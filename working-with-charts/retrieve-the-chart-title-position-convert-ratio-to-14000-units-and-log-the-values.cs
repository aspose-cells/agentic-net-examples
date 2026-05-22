using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

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

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set a title for the chart
        chart.Title.Text = "Sample Chart Title";

        // Retrieve the title position ratios
        double xRatio = chart.Title.XRatioToChart; // fraction of chart width (0‑1)
        double yRatio = chart.Title.YRatioToChart; // fraction of chart height (0‑1)

        // Convert ratios to 1/4000 units (as used by the obsolete X/Y properties)
        int xInUnits = (int)(xRatio * 4000);
        int yInUnits = (int)(yRatio * 4000);

        // Log the values
        Console.WriteLine($"Title XRatioToChart (fraction): {xRatio}");
        Console.WriteLine($"Title YRatioToChart (fraction): {yRatio}");
        Console.WriteLine($"Title X position in 1/4000 units: {xInUnits}");
        Console.WriteLine($"Title Y position in 1/4000 units: {yInUnits}");

        // Save the workbook (optional, just to complete the lifecycle)
        workbook.Save("RetrieveChartTitlePosition.xlsx");
    }
}