// Title: Apply a custom RGB color palette to each series of a column chart using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set individual series colors in a column chart | programmatically change column chart series foreground color with Aspose.Cells | apply custom RGB palette to chart series in Aspose.Cells .NET example | Aspose.Cells chart styling set series Area.ForegroundColor C# | create column chart with specific RGB colors for each series using Aspose.Cells
// Tags: series foreground color Aspose.Cells C# | column chart series color Aspose.Cells | RGB palette chart styling Aspose.Cells | chart series color API Aspose.Cells | Aspose.Cells set series area color

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, defines three RGB colors (red, green, blue), and assigns each color to the corresponding series by setting the series' Area.ForegroundColor property, then saves the file as CustomColorChart.xlsx.
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

            // Populate sample data for three series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Series3");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Custom Color Palette";

            // Add series to the chart (use the data range)
            chart.NSeries.Add("B2:D4", true);

            // Define custom RGB colors for each series
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 0, 0),   // Red for Series1
                Color.FromArgb(0, 255, 0),   // Green for Series2
                Color.FromArgb(0, 0, 255)    // Blue for Series3
            };

            // Apply the custom colors to each series
            for (int i = 0; i < chart.NSeries.Count && i < customColors.Length; i++)
            {
                // For a column chart, set the area foreground color of the series
                chart.NSeries[i].Area.ForegroundColor = customColors[i];
            }

            // Save the workbook with the chart
            workbook.Save("CustomColorChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
