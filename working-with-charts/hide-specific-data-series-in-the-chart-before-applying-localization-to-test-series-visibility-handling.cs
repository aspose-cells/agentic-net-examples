// Title: Hide a Chart Series in Aspose.Cells (C#) Using IsFiltered Before Localization
// Description: Demonstrates how to create a workbook, add a column chart with two data series, hide the second series by setting its IsFiltered property, keep PlotVisibleCellsOnly enabled, and save the file. Ideal for testing chart visibility handling prior to workbook localization.
// Keywords: Aspose.Cells C# hide chart series | IsFiltered property Aspose.Cells | chart series visibility .NET | PlotVisibleCellsOnly Aspose.Cells | column chart hide series | Excel chart filtering Aspose | localization chart series Aspose.Cells | programmatic chart series filter
// Common Searches: C# hide data series in Aspose.Cells chart | Aspose.Cells IsFiltered example | How to filter chart series with Aspose.Cells | PlotVisibleCellsOnly usage Aspose.Cells | Hide second series before localization Aspose
// Developer Intent: Programmatically hide a specific series in an Aspose.Cells chart to control visibility before applying localization or other post‑processing steps.
// Use Cases: Exclude confidential or region‑specific data from a chart that will be localized for different markets. | Allow users to toggle visibility of optional data series in generated Excel reports. | Validate that hidden series do not affect chart rendering when PlotVisibleCellsOnly is active.
// AI Prompts: Generate C# code with Aspose.Cells that hides a selected chart series using IsFiltered and saves the workbook. | Show how to switch the IsFiltered flag for multiple series and confirm PlotVisibleCellsOnly behavior. | Provide a step‑by‑step example of applying localization after hiding certain chart series in an Aspose.Cells workbook.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart with two data series, hide the second series by setting its IsFiltered property, keep PlotVisibleCellsOnly enabled, and save the file. Ideal for testing chart visibility handling prior to workbook localization.
class HideSeriesDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data: categories and two data series
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Series1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Series2");
        worksheet.Cells["C2"].PutValue(15);
        worksheet.Cells["C3"].PutValue(25);
        worksheet.Cells["C4"].PutValue(35);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Add both series to the chart
        chart.NSeries.Add("B2:B4", true); // Series1
        chart.NSeries.Add("C2:C4", true); // Series2
        chart.NSeries.CategoryData = "A2:A4";

        // Hide the second series using the IsFiltered property
        chart.NSeries[1].IsFiltered = true;

        // Ensure only visible cells are plotted (default behavior)
        chart.PlotVisibleCellsOnly = true;

        // Save the workbook
        workbook.Save("HideSeriesDemo.xlsx");
    }
}
