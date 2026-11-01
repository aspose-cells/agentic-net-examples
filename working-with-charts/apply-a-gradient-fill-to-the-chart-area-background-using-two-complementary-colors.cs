// Title: Add a vertical two‑color gradient (blue to orange) to a chart area background with Aspose.Cells for .NET
// AI Prompts: Generate a new Excel workbook, insert a column chart, and apply a vertical blue‑to‑orange gradient to the chart area using Aspose.Cells C# API. | Update an existing chart's ChartArea.FillFormat to use a two‑color linear gradient with custom colors in a .NET application. | Create an Excel file where the chart background is filled with a vertical gradient by configuring FillFormat.SetTwoColorGradient in C#.
// Common Searches: Aspose.Cells C# set chart area background gradient blue orange | How to use FillFormat.SetTwoColorGradient for Excel chart in .NET | vertical linear gradient for chart area with Aspose.Cells example | C# code to apply two‑color gradient to Excel chart background
// Tags: chartarea fillformat two-color gradient aspocells | excel chart background gradient aspocells | set chart area gradient fillformat c# | aspocells chartarea gradient fill | c# gradient fillformat chart area

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart, and uses FillFormat.SetTwoColorGradient to apply a vertical blue‑to‑orange gradient to the chart area's background before saving the file as ChartWithGradientFill.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add sample data for the chart
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a two‑color linear gradient fill to the chart area background (Blue → Orange)
            Color color1 = Color.Blue;
            Color color2 = Color.Orange;
            FillFormat fill = chart.ChartArea.Area.FillFormat;
            // Use Vertical gradient to simulate a linear gradient at 90 degrees
            fill.SetTwoColorGradient(color1, color2, GradientStyleType.Vertical, 90);

            // Save the workbook
            string outputPath = "ChartWithGradientFill.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
