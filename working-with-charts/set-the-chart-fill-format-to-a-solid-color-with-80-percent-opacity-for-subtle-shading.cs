// Title: Set a solid fill with 80% opacity for a chart's plot area using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a column chart and sets the plot area fill to a solid light‑gray color with 20% transparency using Aspose.Cells. | Show how to apply a semi‑transparent solid fill to an Excel chart's plot area with the FillType enum in Aspose.Cells for .NET. | Provide a snippet that adds sample data, inserts a chart, and configures the chart background opacity to 80% in Aspose.Cells.
// Common Searches: aspnet cells how to set chart plot area fill opacity c# | c# aspose.cells set chart background transparency 80 percent | example of solid fill with transparency for Excel chart using Aspose.Cells | apply light gray fill to chart plot area aspose.cells .net
// Tags: Aspose.Cells chart plot area solid fill | C# set chart fill transparency | Aspose.Cells FillType enum usage | Excel chart background opacity .NET | column chart fill format Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Required for FillType enum

// The example creates a workbook, adds sample data, inserts a column chart, and sets the chart's plot area to a light‑gray solid fill with 20 % transparency (80 % opacity) before saving the file as ChartFillOpacity.xlsx.
class SetChartFillOpacity
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart's plot area fill to a solid color with 80% opacity (20% transparency)
            chart.PlotArea.Area.FillFormat.FillType = FillType.Solid;
            chart.PlotArea.Area.FillFormat.SolidFill.Color = Color.LightGray;
            chart.PlotArea.Area.FillFormat.SolidFill.Transparency = 0.2; // 0 = opaque, 1 = fully transparent

            // Save the workbook
            workbook.Save("ChartFillOpacity.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
