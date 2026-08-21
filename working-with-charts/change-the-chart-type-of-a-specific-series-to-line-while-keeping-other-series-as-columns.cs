// Title: Convert a single series to a line chart while retaining column series using Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, add a column chart with two series, switch only the second series to a line type via NSeries.Type, optionally style the line, and save the workbook.
// Keywords: Aspose.Cells C# chart series type | mixed column line chart .NET | change individual series chart type | NSeries.Type property | programmatic Excel chart customization | Aspose.Cells example mixed chart | set series to line chart Aspose
// Common Searches: Aspose.Cells change one series to line chart | mixed column and line chart example C# | how to set chart series type individually Aspose.Cells | convert second series to line in column chart .NET | Aspose.Cells NSeries.Type usage
// Developer Intent: Modify a chart so that only the second data series is displayed as a line while the first series remains a column.
// Use Cases: Quarterly sales report: columns for actual sales, line for sales target. | Financial dashboard: expense categories as columns, cash‑flow trend as a line. | Performance analysis: units sold shown as columns, growth percentage plotted as a line.
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart and changes the second series to a line chart, then saves the workbook. | Explain how the NSeries.Type property enables mixed column‑line charts in Aspose.Cells for .NET. | Provide a step‑by‑step guide to customize the line color of a specific series after converting its chart type.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesTypeDemo
{
    // Shows how to build a workbook, add a column chart with two series, switch only the second series to a line type via NSeries.Type, optionally style the line, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // First series (will stay as column)
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(210);

            // Second series (will be changed to line)
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(90);
            sheet.Cells["C3"].PutValue(130);
            sheet.Cells["C4"].PutValue(160);
            sheet.Cells["C5"].PutValue(190);

            // Add a column chart (default type for all series)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set data range for both series
            chart.NSeries.Add("B2:C5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Change the type of the second series (index 1) to Line
            chart.NSeries[1].Type = ChartType.Line;

            // Optional: customize appearance (e.g., line color)
            chart.NSeries[1].Border.Color = System.Drawing.Color.Red;

            // Save the workbook
            workbook.Save("SeriesTypeChanged.xlsx");
        }
    }
}
