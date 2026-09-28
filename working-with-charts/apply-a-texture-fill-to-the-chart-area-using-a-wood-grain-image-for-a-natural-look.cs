// Title: Apply a wood grain texture fill to a chart area in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a PNG file and uses Aspose.Cells to apply it as a tiled background for a chart's ChartArea. | Show how to configure a chart's FillFormat to use an image texture in Aspose.Cells, including file existence handling. | Provide a step‑by‑step example of creating a column chart in Aspose.Cells and styling its chart area with a custom wood grain image.
// Common Searches: Aspose.Cells C# set chart background image | apply wood grain image to Excel chart area using Aspose.Cells | how to tile a PNG as chart background in Aspose.Cells .NET | example of using FillFormat with image in Aspose.Cells chart
// Tags: chartarea texture fill Aspose.Cells | assign image to chart background Aspose.Cells | Aspose.Cells FillType.Texture usage | C# load PNG for chart texture | chart area tiling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart, loads a wood_grain.png file, assigns its bytes to the chart area's FillFormat as a texture, enables tiling, and saves the workbook as ChartWithWoodTexture.xlsx.
class Program
{
    static void Main()
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

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply a texture fill to the chart area using a wood grain image
        chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

        // Load the wood grain image (ensure the file exists at the specified path)
        string imagePath = "wood_grain.png";
        if (File.Exists(imagePath))
        {
            byte[] imageData = File.ReadAllBytes(imagePath);
            chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;

            // Optionally enable tiling so the texture repeats across the area
            chart.ChartArea.Area.FillFormat.TextureFill.IsTiling = true;
        }
        else
        {
            Console.WriteLine($"Image file not found: {imagePath}");
        }

        // Save the workbook with the textured chart
        workbook.Save("ChartWithWoodTexture.xlsx");
    }
}
