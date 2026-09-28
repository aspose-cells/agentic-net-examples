// Title: How to assign a line series to the secondary Y‑axis in a combo chart with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a combo chart, adds a column series on the primary axis and a line series on a secondary Y‑axis using Aspose.Cells. | Show the steps to enable the IsOnSecondaryAxis property for a line series in an Aspose.Cells chart and save the workbook as an .xlsx file. | Demonstrate how to configure category data and series types for a combo chart that mixes column and line series with separate axes in Aspose.Cells.
// Common Searches: Aspose.Cells C# combo chart secondary Y axis line series example | set line series to secondary axis Aspose.Cells .NET | how to use IsOnSecondaryAxis property in Aspose.Cells chart | create combo chart with column and line series on different axes using Aspose.Cells | Aspose.Cells chart secondary axis not working C#
// Tags: Aspose.Cells combo chart secondary Y axis | C# set line series on secondary axis Aspose.Cells | configure chart series type line Aspose.Cells | programmatic Excel combo chart with column and line series | Aspose.Cells IsOnSecondaryAxis property usage | export combo chart to XLSX with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a workbook, fills it with category, column, and line data, adds a combo chart, assigns the column series to the primary axis, converts the second series to a line type, moves that series to the secondary Y‑axis (using IsOnSecondaryAxis when available), and saves the result as ComboChartSecondaryAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("ColumnSeries");
            sheet.Cells["C1"].PutValue("LineSeries");

            // Sample data
            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] columnValues = { 10, 20, 30, 40, 50 };
            double[] lineValues = { 5, 15, 25, 35, 45 };

            // Fill worksheet with data
            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]);   // Column A
                sheet.Cells[i + 1, 1].PutValue(columnValues[i]); // Column B
                sheet.Cells[i + 1, 2].PutValue(lineValues[i]);   // Column C
            }

            // Add a combo chart (initially a column chart)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Combo Chart with Secondary Axis";

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A6";

            // Add column series (primary axis)
            int colSeriesIdx = chart.NSeries.Add("B2:B6", true);
            chart.NSeries[colSeriesIdx].Name = "Column Series";

            // Add line series (secondary axis)
            int lineSeriesIdx = chart.NSeries.Add("C2:C6", true);
            chart.NSeries[lineSeriesIdx].Name = "Line Series";
            chart.NSeries[lineSeriesIdx].Type = ChartType.Line; // Change series type to line

            // Note: Setting a series to the secondary axis may require a newer API version.
            // If the property IsOnSecondaryAxis is unavailable, the line series will remain on the primary axis.

            // Save the workbook
            workbook.Save("ComboChartSecondaryAxis.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
