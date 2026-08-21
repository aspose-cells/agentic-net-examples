// Title: Add Picture Background to an Aspose.Cells Chart (C#) – FillFormat.TextureFill Example
// Description: Demonstrates how to create a workbook, insert a column chart, set the chart area's FillFormat to Texture, load an image file into a byte array, assign it to TextureFill.ImageData, and save the Excel file. Includes error handling for missing images.
// Keywords: Aspose.Cells chart background image | C# FillFormat TextureFill | Excel chart area picture fill | .NET Aspose.Cells example | load image bytes into chart | chart area texture fill | Aspose.Cells chart styling
// Common Searches: Aspose.Cells set chart background picture C# | FillFormat TextureFill chart example .NET | how to add image to chart area Aspose.Cells | C# load PNG into Excel chart background | Aspose.Cells chart area fill from file
// Developer Intent: Apply an image file as the background of a chart by using FillFormat.TextureFill in Aspose.Cells for .NET.
// Use Cases: Generate sales dashboards where each chart displays a company logo as a background. | Create branded marketing reports with themed picture fills for multiple charts. | Automate recurring financial presentations that require a consistent chart background across workbooks.
// AI Prompts: Show how to modify the snippet to use a JPEG image and stretch it to cover the entire chart area. | Provide code that assigns different picture backgrounds to several charts in the same workbook. | Explain a fallback strategy that applies a solid color fill when the image file cannot be found.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, insert a column chart, set the chart area's FillFormat to Texture, load an image file into a byte array, assign it to TextureFill.ImageData, and save the Excel file. Includes error handling for missing images.
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

            // Set the chart area fill type to texture (picture background)
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            // Load the image file into a byte array if it exists
            const string imagePath = "background.png";
            if (File.Exists(imagePath))
            {
                byte[] imageData = File.ReadAllBytes(imagePath);
                // Apply the image data to the chart area's texture fill
                chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Chart background will not be set.");
            }

            // Save the workbook with the chart background applied
            const string outputPath = "ChartWithBackground.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
