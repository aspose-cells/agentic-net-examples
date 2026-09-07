// Title: How to apply a PNG picture as a stretched background fill for a chart area using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a PNG file and assigns it to a chart's ChartArea.FillFormat as a texture fill with Aspose.Cells. | Show how to configure the FillPictureType to Stretch so the image covers the entire chart area in a column chart. | Create a robust snippet that checks for the image file, applies it as the chart background, and saves the workbook while handling missing‑file warnings.
// Common Searches: Aspose.Cells set chart area background image from file in C# | C# load PNG into FillFormat texture fill for Aspose.Cells chart | How to stretch a picture fill to cover a chart area using Aspose.Cells .NET | Apply background picture to column chart with Aspose.Cells and handle missing image
// Tags: chartarea texture fill Aspose.Cells | set chart background image C# | load PNG into FillFormat Aspose.Cells | stretch picture fill chart area | column chart background image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a column chart, sets the chart area fill type to texture, loads a PNG file named 'background.png' (if present), assigns the image data to the chart's texture fill, stretches the picture to fill the chart area, and saves the workbook as 'ChartWithBackground.xlsx', with a warning when the image file is missing.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 5, 15, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the chart area fill type to texture (picture)
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            // Load the background image if it exists
            string imagePath = "background.png";
            if (File.Exists(imagePath))
            {
                byte[] imageData = File.ReadAllBytes(imagePath);
                // Apply the image as the texture fill for the chart area
                chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;
                // Stretch the picture to fill the chart area
                chart.ChartArea.Area.FillFormat.PictureFormatType = FillPictureType.Stretch;
            }
            else
            {
                Console.WriteLine($"Warning: Image file '{imagePath}' not found. Chart will use default background.");
            }

            // Save the workbook with the chart background applied
            workbook.Save("ChartWithBackground.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
