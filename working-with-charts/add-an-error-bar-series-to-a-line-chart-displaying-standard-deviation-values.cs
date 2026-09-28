// Title: Create a line chart with standard deviation Y‑error bars using Aspose.Cells for .NET
// AI Prompts: Generate a line chart from worksheet data and attach Y‑error bars that use the series' standard deviation as the error amount with Aspose.Cells. | Set the series YErrorBar.Type to StDev and configure the DisplayType to show both positive and negative error bars. | Save the workbook as an .xlsx file after adding the error bars.
// Common Searches: how to show error bars based on standard deviation in an Aspose.Cells line chart C# | setting YErrorBar type to StDev for a series using Aspose.Cells .NET | displaying both positive and negative error bars in a chart created with Aspose.Cells | adding error bars to a line chart and saving as XLSX with Aspose.Cells
// Tags: Aspose.Cells line chart YErrorBar StDev | C# Aspose.Cells chart series error bars | add error bars to line chart .NET | save workbook with chart error bars | configure error bar display both directions

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a workbook, populating X/Y data, adding a line chart, applying standard deviation Y‑error bars (both directions) to the series, and saving the file as LineChartWithStdDevErrorBar.xlsx.
class LineChartWithStdDevErrorBar
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for X (categories) and Y (values)
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["A5"].PutValue(4);

        sheet.Cells["B1"].PutValue("Y");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(40);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 6, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add a series: Y values are taken from column B, X categories from column A
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Configure the Y‑direction error bar to display standard deviation
        Series series = chart.NSeries[0];
        series.YErrorBar.Type = ErrorBarType.StDev;               // Use standard deviation as error amount
        series.YErrorBar.DisplayType = ErrorBarDisplayType.Both; // Show both plus and minus error bars
        // When Type is StDev, the Amount property is not required

        // Save the workbook (lifecycle: save)
        workbook.Save("LineChartWithStdDevErrorBar.xlsx");
    }
}
