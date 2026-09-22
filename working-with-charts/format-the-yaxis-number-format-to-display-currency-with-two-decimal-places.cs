// Title: Format the Y‑axis of an Aspose.Cells column chart as currency with two decimal places in C#
// AI Prompts: Write C# code that creates a workbook, adds numeric data, applies the "$#,##0.00" number format to the data range, and ensures the column chart Y‑axis shows values as currency with two decimal places using Aspose.Cells. | Show how to set a custom number format on worksheet cells so that an Aspose.Cells chart automatically inherits a currency format on its Y‑axis. | Provide a complete example that builds a column chart from revenue data and configures the axis number format to display dollars and cents in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# display chart Y axis in dollars with cents | Set currency format for Excel chart axis using Aspose.Cells .NET | How to make Y‑axis of column chart show $ values in Aspose.Cells | Apply number format to chart data range so axis shows $ in Aspose.Cells | C# Aspose.Cells chart axis formatting for monetary values
// Tags: currency number format chart axis Aspose.Cells | apply number format to worksheet cells C# | column chart Y axis formatting Aspose.Cells .NET | inherit cell style in chart axis Aspose.Cells | format revenue values as dollars in Excel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, inserts month and revenue data, applies a "$#,##0.00" currency style to the revenue cells, adds a column chart that uses those cells as its series, and saves the file. Because the cells are formatted as currency, the chart's Y‑axis automatically displays values with a dollar sign and two decimal places.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Get the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the chart
            ws.Cells["A1"].PutValue("Month");
            ws.Cells["B1"].PutValue("Revenue");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["B2"].PutValue(1234.56);
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["B3"].PutValue(7890.12);

            // Apply currency format to the revenue cells (so the axis inherits it)
            Style currencyStyle = wb.CreateStyle();
            currencyStyle.Custom = "$#,##0.00";
            ws.Cells.CreateRange("B2:B3").ApplyStyle(currencyStyle, new StyleFlag { NumberFormat = true });

            // Add a column chart to the worksheet
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Save the workbook to a file
            string outputPath = "ChartWithCurrencyAxis.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
