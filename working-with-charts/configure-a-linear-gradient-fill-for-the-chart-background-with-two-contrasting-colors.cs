// Title: Create a horizontal two‑color gradient background for an Excel chart using Aspose.Cells for .NET (C#)
// Common Searches: how to add a horizontal two color gradient to a chart area with Aspose.Cells C# | Aspose.Cells SetTwoColorGradient example for chart background | C# code to apply linear gradient fill to an Excel chart using Aspose.Cells | set chart area fill format gradient style horizontal Aspose.Cells .NET
// Tags: SetTwoColorGradient chart area Aspose.Cells | horizontal gradient fill Excel chart .NET | chart background linear gradient Aspose.Cells | Aspose.Cells chart area styling C# | Excel chart gradient background using Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The sample creates a workbook, adds sample data, inserts a column chart, and then uses FillFormat.SetTwoColorGradient with GradientStyleType.Horizontal to apply a blue‑to‑yellow horizontal gradient to the chart's background before saving the file as ChartWithGradient.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            var chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a horizontal two‑color gradient fill to the chart background
            var fill = chart.ChartArea.Area.FillFormat;
            // Use GradientStyleType.Horizontal for compatibility with the current Aspose.Cells version
            fill.SetTwoColorGradient(Color.Blue, Color.Yellow, GradientStyleType.Horizontal, 0);

            // Save the workbook with the chart
            workbook.Save("ChartWithGradient.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
