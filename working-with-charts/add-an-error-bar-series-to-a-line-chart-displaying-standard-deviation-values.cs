// Title: Add Standard Deviation Y‑Error Bars to a Line Chart with Aspose.Cells for .NET
// Description: Creates a workbook, inserts X/Y data, adds a line chart, binds the series, and configures the Y‑error bar to show standard‑deviation values with both positive and negative bars before saving as an .xlsx file.
// Keywords: Aspose.Cells | C# chart error bars | line chart error bar | YErrorBar StDev | ErrorBarDisplayType Both | standard deviation error bar | .NET Excel chart | Aspose.Cells chart customization
// Common Searches: Aspose.Cells line chart standard deviation error bars | C# add Y error bar StDev Aspose.Cells | show both plus and minus error bars in Aspose.Cells chart | how to configure error bars for a series using Aspose.Cells .NET | add error bars to Excel chart programmatically
// Developer Intent: Add a Y‑error bar that displays standard‑deviation values to a line chart created with Aspose.Cells.
// Use Cases: Visualize measurement variability in scientific data by adding standard‑deviation error bars to line charts. | Generate automated reports where each series includes confidence intervals for quality control. | Build dashboards that require both positive and negative error bars to illustrate statistical uncertainty.
// AI Prompts: Show how to add X‑error bars with custom values to the same chart using Aspose.Cells. | Provide code to change the color and thickness of standard‑deviation Y‑error bars in an Aspose.Cells chart. | Explain how to retrieve and modify error‑bar settings for multiple series in an Aspose.Cells workbook.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, inserts X/Y data, adds a line chart, binds the series, and configures the Y‑error bar to show standard‑deviation values with both positive and negative bars before saving as an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the line chart
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["A5"].PutValue(4);

        sheet.Cells["B1"].PutValue("Y");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(12);
        sheet.Cells["B4"].PutValue(14);
        sheet.Cells["B5"].PutValue(16);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 6, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add the series data and category (X‑axis) data
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Configure the Y‑error bar to display standard deviation values
        Series series = chart.NSeries[0];
        series.YErrorBar.Type = ErrorBarType.StDev;               // Use standard deviation
        series.YErrorBar.DisplayType = ErrorBarDisplayType.Both; // Show both plus and minus bars

        // Save the workbook with the chart
        workbook.Save("LineChartWithStdDevErrorBar.xlsx");
    }
}
