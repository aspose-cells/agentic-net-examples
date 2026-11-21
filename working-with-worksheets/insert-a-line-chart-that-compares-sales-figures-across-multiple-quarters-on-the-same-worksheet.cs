// Title: Create and insert a multi-series line chart for quarterly sales data into an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to add a multi-series line chart to an Excel worksheet with Aspose.Cells C# | Aspose.Cells example for plotting quarterly sales comparison chart | set category axis range for a line chart in Aspose.Cells .NET | position chart legend on the right using Aspose.Cells C# | specify chart location by rows and columns in Aspose.Cells
// Tags: add multi-series line chart Aspose.Cells | set category axis range Aspose.Cells | position chart legend right Aspose.Cells | define chart placement rows columns Aspose.Cells | quarterly sales line chart .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills cells with quarter labels and sales figures for two products, adds a line chart positioned between rows 7‑25 and columns 0‑10, sets the chart title, adds two data series referencing the product ranges, assigns the quarter range as the category axis, moves the legend to the right, and saves the file as 'QuarterlySalesLineChart.xlsx'.
class LineChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data: quarters and sales for two products
            sheet.Cells["A1"].PutValue("Quarter");
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["C1"].PutValue("Product B");

            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            sheet.Cells["B2"].PutValue(12000);
            sheet.Cells["B3"].PutValue(15000);
            sheet.Cells["B4"].PutValue(13000);
            sheet.Cells["B5"].PutValue(17000);

            sheet.Cells["C2"].PutValue(10000);
            sheet.Cells["C3"].PutValue(14000);
            sheet.Cells["C4"].PutValue(12000);
            sheet.Cells["C5"].PutValue(16000);

            // Add a line chart to the worksheet
            // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
            int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Quarterly Sales Comparison";

            // Add series for Product A
            int seriesIndexA = chart.NSeries.Add("B2:B5", true);
            chart.NSeries[seriesIndexA].Name = "Product A";

            // Add series for Product B
            int seriesIndexB = chart.NSeries.Add("C2:C5", true);
            chart.NSeries[seriesIndexB].Name = "Product B";

            // Set category (X) axis labels to quarters
            chart.NSeries.CategoryData = "A2:A5";

            // Position legend on the right
            chart.Legend.Position = LegendPositionType.Right;

            // Save the workbook to a file
            string outputPath = "QuarterlySalesLineChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
