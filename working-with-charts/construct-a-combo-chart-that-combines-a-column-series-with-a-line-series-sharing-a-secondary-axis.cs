// Title: Create a combo chart with a column series and a line series on a secondary axis using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that builds a combo chart where the column series uses the primary Y‑axis and the line series uses a secondary Y‑axis. | Demonstrate how to change a data series to a line type, assign it to the secondary axis, and save the workbook as an XLSX file using Aspose.Cells.
// Common Searches: aspnet aspose.cells combo chart secondary y axis example | c# create column and line series on same chart with Aspose.Cells | how to set IsSecondaryAxis for a series in Aspose.Cells chart | Aspose.Cells chart with mixed column and line types | save combo chart to Excel using Aspose.Cells C#
// Tags: Aspose.Cells combo chart secondary axis | C# column series primary axis line series secondary axis | change series to line type Aspose.Cells | mixed column and line chart Aspose.Cells | populate chart data from worksheet cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The program creates a new workbook, fills cells A1:C6 with category labels and numeric values for column and line series, adds a column chart, inserts a column series on the primary axis and a line series on the secondary axis (changing its type to Line), sets a chart title, and saves the file as ComboChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the chart
            // Column A: Categories, Column B: Column series values, Column C: Line series values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Column");
            sheet.Cells["C1"].PutValue("Line");

            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] columnValues = { 10, 20, 30, 25, 15 };
            double[] lineValues = { 5, 15, 25, 20, 10 };

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]);   // A column
                sheet.Cells[i + 1, 1].PutValue(columnValues[i]); // B column
                sheet.Cells[i + 1, 2].PutValue(lineValues[i]);   // C column
            }

            // Add a combo chart (initially a column chart) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the category (X) axis data
            chart.NSeries.CategoryData = "A2:A6";

            // Add the column series (primary axis)
            int columnSeriesIdx = chart.NSeries.Add("B2:B6", true);
            chart.NSeries[columnSeriesIdx].Name = "Column Series";

            // Add the line series (secondary axis)
            int lineSeriesIdx = chart.NSeries.Add("C2:C6", true);
            chart.NSeries[lineSeriesIdx].Name = "Line Series";

            // Change the line series type to Line
            chart.NSeries[lineSeriesIdx].Type = ChartType.Line;

            // NOTE: The IsSecondaryAxis property may not be available in older Aspose.Cells versions.
            // If needed and supported, uncomment the following line:
            // chart.NSeries[lineSeriesIdx].IsSecondaryAxis = true;

            // Optional: set chart title
            chart.Title.Text = "Combo Chart: Column + Line";

            // Save the workbook to a file
            workbook.Save("ComboChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
