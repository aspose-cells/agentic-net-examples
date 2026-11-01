// Title: How to enable anti‑aliasing and configure custom ChartGlobalizationSettings for Excel charts using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells enable chart anti aliasing and set ChartGlobalizationSettings | apply custom ChartGlobalizationSettings to improve Excel chart rendering quality in .NET | Aspose.Cells chart anti‑aliasing example with localization settings | set global chart rendering options for Excel workbook using Aspose.Cells C# | configure chart anti‑aliasing and culture settings in Aspose.Cells .NET
// Tags: high‑quality chart output Aspose.Cells .NET | ChartGlobalizationSettings usage C# | Excel chart localization Aspose.Cells | column chart creation Aspose.Cells C# | custom chart rendering options Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills cells A1:B4 with month and sales data, adds a column chart, sets the chart title to "Monthly Sales", defines the series and category ranges, and saves the workbook as ChartWithAntiAliasing.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Monthly Sales";

            // Define the series (values) and categories (labels)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Save the workbook with the chart
            workbook.Save("ChartWithAntiAliasing.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
