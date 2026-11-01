// Title: Define a custom RGB color palette and apply distinct colors to each series of an Aspose.Cells column chart in C#
// AI Prompts: Write C# code using Aspose.Cells to create a column chart and set a unique RGB fill color for each series. | Show how to assign a black border and custom foreground color to chart series in an Aspose.Cells workbook. | Update an existing Aspose.Cells chart by applying a predefined Color[] palette to its series and save the workbook.
// Common Searches: Aspose.Cells C# set individual series colors in a column chart | How to apply custom RGB palette to chart series with Aspose.Cells .NET | Programmatically change fill and border colors of Excel chart series using Aspose.Cells | Assign different colors to each series in an Aspose.Cells generated chart
// Tags: Aspose.Cells series color fill | Aspose.Cells chart series palette | Excel column chart series styling .NET | chart series border styling Aspose.Cells | apply series colors Aspose.Cells C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, defines a three‑color RGB palette (red, green, blue), assigns each color to the matching series, sets a black border for visibility, and saves the file as CustomPaletteChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the chart
            ws.Cells["A1"].PutValue("Month");
            ws.Cells["B1"].PutValue("Series1");
            ws.Cells["C1"].PutValue("Series2");
            ws.Cells["D1"].PutValue("Series3");

            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");

            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);

            ws.Cells["C2"].PutValue(15);
            ws.Cells["C3"].PutValue(25);
            ws.Cells["C4"].PutValue(35);

            ws.Cells["D2"].PutValue(12);
            ws.Cells["D3"].PutValue(22);
            ws.Cells["D4"].PutValue(32);

            // Add a column chart (type, upper-left row, upper-left column, lower-right row, lower-right column)
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the chart (including categories)
            chart.NSeries.Add("B2:D4", true);

            // Define a custom color palette for the series
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 0, 0),   // Red for Series1
                Color.FromArgb(0, 255, 0),   // Green for Series2
                Color.FromArgb(0, 0, 255)    // Blue for Series3
            };

            // Assign each custom color to the corresponding series
            for (int i = 0; i < chart.NSeries.Count && i < customColors.Length; i++)
            {
                // Set the fill color of the series
                chart.NSeries[i].Area.ForegroundColor = customColors[i];

                // Set a border color for better visibility
                chart.NSeries[i].Border.IsVisible = true;
                chart.NSeries[i].Border.Color = Color.Black;
            }

            // Save the workbook with the chart
            string outputPath = "CustomPaletteChart.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
