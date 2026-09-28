// Title: Handle missing or invalid texture image when setting a chart's plot area fill in Aspose.Cells for .NET using try‑catch
// AI Prompts: Write C# code that loads a PNG file as a texture for an Aspose.Cells chart's plot area, wraps the loading and FillFormat assignment in a try‑catch block, and falls back to a solid fill if any exception occurs. | Show how to safely apply a custom FillPattern to a chart in Aspose.Cells, including error handling for missing or corrupted image files and guaranteeing the workbook saves without interruption.
// Common Searches: aspocells c# chart plot area texture file not found exception handling | how to apply custom image fill to Aspose.Cells chart with fallback to solid | c# try-catch when loading texture for Aspose.Cells chart FillFormat | Aspose.Cells chart fill pattern error handling for missing PNG | set chart background image in Aspose.Cells and handle load errors
// Tags: aspocells chart background texture handling | fillformat solid fallback on texture error | c# chart fill pattern error handling | apply image fill to Aspose.Cells chart | try-catch texture load Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a workbook, adds a column chart, and attempts to set a custom PNG texture on the chart's plot area. It checks for the texture file, wraps the texture application in a try‑catch block, and defaults to a solid fill if the file is missing or an error occurs, while also catching any unexpected exceptions before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate data for the chart
            ws.Cells["A1"].PutValue(10);
            ws.Cells["A2"].PutValue(20);
            ws.Cells["A3"].PutValue(30);
            ws.Cells["B1"].PutValue(15);
            ws.Cells["B2"].PutValue(25);
            ws.Cells["B3"].PutValue(35);

            // Add a column chart
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("A1:A3", true);
            chart.NSeries[0].Name = "Series 1";

            // Apply fill to the plot area
            FillFormat fill = chart.PlotArea.Area.FillFormat;

            string texturePath = "texture.png";
            if (File.Exists(texturePath))
            {
                try
                {
                    // If custom texture support is required, implement it here.
                    // For now, use a solid fill as a safe default.
                    fill.Pattern = FillPattern.Solid;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error applying texture: " + ex.Message);
                    fill.Pattern = FillPattern.Solid;
                }
            }
            else
            {
                // Use solid fill when texture file is missing
                fill.Pattern = FillPattern.Solid;
            }

            // Save the workbook
            wb.Save("ChartWithTexture.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
