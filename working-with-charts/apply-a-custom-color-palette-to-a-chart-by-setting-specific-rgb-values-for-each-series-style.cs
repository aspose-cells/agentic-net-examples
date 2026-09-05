// Title: Apply a custom RGB color palette to each series of a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates an Excel workbook, adds a column chart, and assigns specific RGB values to the foreground color of each series with Aspose.Cells. | Show how to programmatically set the ForegroundColor of chart series in Aspose.Cells to customize a column chart's palette. | Provide an example that populates sample data, generates a column chart, and applies red, green, and blue colors to three series before saving the file.
// Common Searches: Aspose.Cells C# set custom RGB colors for individual chart series | How to change series colors in an Excel column chart using Aspose.Cells .NET | Example of applying a custom color palette to a chart with Aspose.Cells for .NET | Programmatic way to assign different colors to each series in Aspose.Cells chart | Set chart series foreground color Aspose.Cells C# tutorial
// Tags: set chart series RGB color Aspose.Cells | custom column chart series colors .NET | Aspose.Cells chart series foreground color | apply custom palette to Excel chart C# | programmatic chart series styling Aspose.Cells | color each series in Aspose.Cells column chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// The program creates a workbook, fills it with sample data, adds a column chart, and assigns red, green, and blue RGB colors to the three series by setting each series' ForegroundColor, then saves the file as CustomPaletteChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Series3");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:D4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define custom RGB colors for each series
            Color[] seriesColors = new Color[]
            {
                Color.FromArgb(255, 0, 0),   // Red for Series1
                Color.FromArgb(0, 255, 0),   // Green for Series2
                Color.FromArgb(0, 0, 255)    // Blue for Series3
            };

            // Apply the custom colors to the series
            for (int i = 0; i < chart.NSeries.Count && i < seriesColors.Length; i++)
            {
                chart.NSeries[i].Area.ForegroundColor = seriesColors[i];
            }

            // Save the workbook with the chart
            workbook.Save("CustomPaletteChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
