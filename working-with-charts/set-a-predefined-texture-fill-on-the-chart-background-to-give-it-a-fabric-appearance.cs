// Title: How to apply a fabric texture to an Excel chart background with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart area texture fill using image bytes | Add fabric background to Excel chart with Aspose.Cells .NET | C# example for applying texture fill to chart background in Aspose.Cells | Load PNG as texture for chart area in Aspose.Cells workbook | Set chart background opacity and border after texture fill Aspose.Cells
// Tags: chart area texture fill Aspose.Cells | apply image texture to Excel chart background C# | FillFormat.FillType.Texture Aspose.Cells example | chart background border styling Aspose.Cells | load image bytes for chart fill Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace ChartTextureFillDemo
{
    // The sample loads a fabric.png file into a byte array, creates a workbook with sample data, adds a column chart, and sets the chart area's FillFormat.FillType to Texture using the image bytes. It ensures full opacity, adds a thin black border around the chart area, and saves the workbook as ChartWithFabricTexture.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the texture image (fabric appearance). Adjust if necessary.
            const string textureImagePath = "fabric.png";

            // Verify that the texture image exists.
            if (!File.Exists(textureImagePath))
            {
                Console.WriteLine($"Error: Texture image file \"{textureImagePath}\" not found.");
                return;
            }

            try
            {
                // Load the texture image bytes.
                byte[] textureBytes = File.ReadAllBytes(textureImagePath);

                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Get the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data for the chart.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B4"].PutValue(180);

                // Add a column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart.
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -----------------------------------------------------------------
                // Apply a texture fill to the chart background (ChartArea).
                // -----------------------------------------------------------------
                // Set the fill type to Texture.
                chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

                // Assign the texture image data.
                chart.ChartArea.Area.FillFormat.TextureFill.ImageData = textureBytes;

                // Ensure the chart background is visible (no transparency).
                chart.ChartArea.Area.FillFormat.Transparency = 0.0;

                // -----------------------------------------------------------------
                // (Optional) Add a thin border around the chart area to highlight the texture.
                // -----------------------------------------------------------------
                chart.ChartArea.Border.Color = Color.Black;
                // Use the WeightType enum for line weight.
                chart.ChartArea.Border.Weight = WeightType.SingleLine;

                // Save the workbook.
                const string outputFile = "ChartWithFabricTexture.xlsx";
                workbook.Save(outputFile, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully as \"{outputFile}\".");
                Console.WriteLine("The chart background now uses the specified fabric texture.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart with texture fill:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
