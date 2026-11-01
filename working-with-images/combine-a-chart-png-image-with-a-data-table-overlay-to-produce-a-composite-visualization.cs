// Title: Create a column chart from worksheet data, add a data table overlay, and export the result as a PNG image using Aspose.Cells for .NET
// AI Prompts: Write C# code that builds a workbook, fills it with sample data, creates a column chart, adds the chart’s data table, and saves the combined visualization as a PNG file with Aspose.Cells. | Generate a C# snippet that resizes the chart, positions a data table beneath it, and exports the chart‑plus‑table image to a PNG using Aspose.Cells. | Provide C# example code that extracts a chart image, renders the associated data table as a bitmap, merges the two graphics, and writes the composite PNG to disk with Aspose.Cells.
// Common Searches: Aspose.Cells C# add data table to chart before exporting to PNG | how to overlay a data table on an Excel chart image using Aspose.Cells .NET | export column chart with data table as PNG with Aspose.Cells | C# Aspose.Cells combine chart image and data table into one PNG file | save Excel chart with data table overlay as image using Aspose.Cells
// Tags: export chart with data table to PNG Aspose.Cells | column chart generation from worksheet data C# | add data table overlay to chart Aspose.Cells | chart image creation Aspose.Cells .NET | save Excel chart as PNG with data table

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes category and value data to cells A1:B5, builds a column chart linked to that range, sets a chart title, and saves the chart as a PNG file named 'CompositeChart.png' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create workbook and populate data ----------
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Header
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Sample data
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(25);

            // ---------- Add a column chart ----------
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            var chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B5", true);               // Values
            chart.NSeries.CategoryData = "A2:A5";           // Categories
            chart.Title.Text = "Sample Column Chart";

            // ---------- Export chart to PNG ----------
            // The ToImage method without specifying format saves as PNG by default.
            chart.ToImage("CompositeChart.png");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
