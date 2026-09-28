// Title: How to Verify an Image File Exists Before Using It as a Chart Background Texture in Aspose.Cells for .NET
// AI Prompts: Check for the presence of a PNG file and apply it as a texture fill to a chart's background using Aspose.Cells in C#. | Add error handling that throws a FileNotFoundException when the specified chart background image cannot be located. | Read an image into a byte array and assign the data to ChartArea.Area.FillFormat.TextureFill.ImageData in a workbook. | Save the workbook after setting the chart's background texture, ensuring the image file was validated first.
// Common Searches: aspocells c# set chart background image only if file exists | how to add texture fill to chart area with image validation in Aspose.Cells | c# Aspose.Cells chart background texture file not found handling | example of using File.Exists before assigning chart background image in Aspose.Cells
// Tags: chart background texture fill Aspose.Cells | validate image file existence C# | TextureFill.ImageData assignment Aspose.Cells | File.Exists error handling Aspose.Cells chart | apply PNG as chart area texture .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds a column chart, sets the chart area fill type to Texture, verifies that 'chart_background.png' exists, reads the image into a byte array, assigns the bytes to the chart's TextureFill.ImageData, and saves the workbook as 'ChartWithBackgroundImage.xlsx' with proper error handling for missing files and save failures.
    public class ChartBackgroundImageValidation
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a sample column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 5, 15, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart area fill type to Texture so we can apply an image
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            // Path to the image that will be used as chart background
            string imagePath = "chart_background.png";

            // Validate that the image file exists before loading it
            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException($"Image file not found: {imagePath}");
            }

            // Read the image file into a byte array
            byte[] imageData = File.ReadAllBytes(imagePath);

            // Assign the image data to the texture fill of the chart background
            chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;

            // Optional: set additional fill properties (e.g., stretch the texture)
            // The default fill type is Stretch; if needed, uncomment the line below
            // chart.ChartArea.Area.FillFormat.TextureFill.Type = TextureFillType.Stretch;

            // Save the workbook with the chart that now has a background image
            string outputPath = "ChartWithBackgroundImage.xlsx";

            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
