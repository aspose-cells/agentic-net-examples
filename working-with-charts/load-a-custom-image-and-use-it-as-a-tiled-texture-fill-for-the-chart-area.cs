// Title: How to fill a chart area with a tiled custom image texture using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart area background to a repeating image texture | Load PNG into byte array and apply as tiled texture fill for a chart in .NET | Create column chart with custom tiled texture background using Aspose.Cells | Enable IsTiling on chart area texture fill in C# Aspose.Cells example | Apply texture fill to chart area from an image file in Aspose.Cells for .NET
// Tags: chart area texture fill Aspose.Cells | custom image tiled background .NET | FillType.Texture usage | TextureFill ImageData assignment | IsTiling property chart background

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example loads a PNG file into a byte array, creates a workbook with sample data, adds a column chart, and configures the chart area to use a texture fill. It sets FillType to Texture, assigns the image data to the TextureFill object, enables tiling via IsTiling, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Path to the custom image that will be used as a tiled texture.
        // Ensure that the file exists in the same folder as the executable or provide a full path.
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "texture.png");

        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Image file not found: {imagePath}");
            return;
        }

        try
        {
            // Load image data into a byte array.
            byte[] imageData = File.ReadAllBytes(imagePath);

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart.
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the chart area to use a tiled texture fill.
            // 1. Set the fill type to Texture.
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            // 2. Access the TextureFill object.
            TextureFill textureFill = chart.ChartArea.Area.FillFormat.TextureFill;

            // 3. Assign the custom image data.
            textureFill.ImageData = imageData;

            // 4. Enable tiling so the image repeats across the chart area.
            textureFill.IsTiling = true;

            // Optional: adjust scaling (0.0 – 1.0). Here we keep the original size.
            textureFill.Scale = 1.0;

            // Save the workbook.
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ChartWithTiledTexture.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the chart with tiled texture:");
            Console.WriteLine(ex.Message);
        }
    }
}
