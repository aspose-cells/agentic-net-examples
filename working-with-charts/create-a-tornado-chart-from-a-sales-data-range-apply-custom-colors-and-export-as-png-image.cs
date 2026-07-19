// Title: Create a Tornado (Stacked Horizontal Bar) Chart with Custom Colors and Export to PNG using Aspose.Cells for .NET
// Description: This example builds a new workbook, fills it with region‑wise sales figures for two products, adds a stacked horizontal bar chart (tornado style), applies distinct fill colors to each series, and saves the chart as a PNG image while optionally keeping the Excel file.
// Keywords: Aspose.Cells tornado chart C# | stacked horizontal bar chart Aspose.Cells | export chart to PNG Aspose.Cells | custom series colors Aspose.Cells | C# generate chart image | sales data visualization Aspose.Cells | chart image output .NET | Aspose.Cells chart formatting
// Common Searches: how to make a tornado chart with Aspose.Cells | set series fill color Aspose.Cells C# | export Aspose.Cells chart as PNG | stacked bar chart from range Aspose.Cells | C# code for horizontal bar chart image
// Developer Intent: Generate a tornado chart from sales data, color each series uniquely, and save the chart as a PNG file using Aspose.Cells for .NET.
// Use Cases: Compare product sales across regions in a management presentation. | Automate creation of chart images for web dashboards without opening Excel. | Provide both a visual PNG chart and the source workbook for stakeholder distribution.
// AI Prompts: Write C# code with Aspose.Cells to create a tornado chart from a given range, apply custom colors to the series, and export the chart as a PNG file. | Explain how to change fill colors of chart series and set the chart type to stacked horizontal bar in Aspose.Cells for .NET. | Show the steps to save a chart created with Aspose.Cells as an image while also preserving the workbook.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example builds a new workbook, fills it with region‑wise sales figures for two products, adds a stacked horizontal bar chart (tornado style), applies distinct fill colors to each series, and saves the chart as a PNG image while optionally keeping the Excel file.
class TornadoChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample sales data
            // Column A: Region, Column B: Sales for Product A, Column C: Sales for Product B
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["C1"].PutValue("Product B");

            string[] regions = { "North", "South", "East", "West", "Central" };
            int[] salesA = { 120, 80, 150, 70, 110 };
            int[] salesB = { 100, 90, 130, 60, 95 };

            for (int i = 0; i < regions.Length; i++)
            {
                int row = i + 2; // Data starts from row 2
                sheet.Cells[row, 0].PutValue(regions[i]); // Column A
                sheet.Cells[row, 1].PutValue(salesA[i]);  // Column B
                sheet.Cells[row, 2].PutValue(salesB[i]);  // Column C
            }

            // Add a stacked bar chart (horizontal) which will serve as a tornado chart
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 8, 1, 25, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the first series (Product A)
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries[0].Name = "Product A";

            // Add the second series (Product B)
            chart.NSeries.Add("C2:C6", true);
            chart.NSeries[1].Name = "Product B";

            // Apply custom colors to each series
            chart.NSeries[0].Area.FillFormat.SolidFill.Color = Color.CornflowerBlue;
            chart.NSeries[1].Area.FillFormat.SolidFill.Color = Color.OrangeRed;

            // Export the chart as a PNG image
            string chartImagePath = "tornado_chart.png";
            chart.ToImage(chartImagePath);

            // Save the workbook (optional, just to keep the Excel file)
            string workbookPath = "tornado_chart.xlsx";
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
