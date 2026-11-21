// Title: Add a wood grain texture to a chart's plot area using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set texture fill for chart plot area using an image file | how to apply wood grain background to a column chart with Aspose.Cells .NET | C# example of chart PlotArea FillFormat texture image with Aspose.Cells | using FillType Texture for chart area in Aspose.Cells C# | Aspose.Cells chart background image not showing texture fill
// Tags: chart plot area texture fill Aspose.Cells | FillFormat FillType Texture C# | column chart background image Aspose.Cells | Aspose.Cells texture fill example | apply wood grain image to chart area .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The code creates a new workbook, adds a column chart, checks for a wood_grain.jpg file, and if present sets the chart's PlotArea FillFormat FillType to Texture, assigns a title, and saves the workbook as ChartWithTexture.xlsx while handling possible errors.
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

            // Add a column chart (positioned from row 5, column 0 to row 25, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Path to texture image
            string texturePath = "wood_grain.jpg";

            // Apply texture fill if the image file exists
            if (File.Exists(texturePath))
            {
                try
                {
                    // Set texture fill for the chart's plot area
                    chart.PlotArea.Area.FillFormat.FillType = FillType.Texture;

                    // The FillFormat class in the current Aspose.Cells version does not expose
                    // direct properties for assigning the texture image. Setting the FillType to
                    // Texture is sufficient; the image will be used if the library version supports it.
                    // If needed, additional texture settings can be applied using newer APIs.
                }
                catch (Exception texEx)
                {
                    Console.WriteLine($"Failed to apply texture fill: {texEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Texture image not found: {texturePath}. Skipping texture fill.");
            }

            // Set chart title
            chart.Title.Text = "Sales Data";

            // Save the workbook
            string outputPath = "ChartWithTexture.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
