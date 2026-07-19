// Title: Aspose.Cells for .NET – Apply a File‑Based Image as Chart Background (C#)
// Description: This C# sample creates a workbook, adds a column chart, switches the chart area fill to a texture, reads a PNG/JPEG logo from a given path into a byte array, assigns the bytes to the chart’s TextureFill, and saves the file while gracefully handling a missing image.
// Keywords: Aspose.Cells C# chart background image | Excel chart texture fill | Load image file into chart area | Set chart area FillFormat TextureFill | company logo watermark Excel | chart background from file path | Aspose.Cells FillType.Texture | chart area image data | C# Aspose.Cells example
// Common Searches: how to set a chart background image using Aspose.Cells .NET | Aspose.Cells chart area texture fill from file | C# add logo to Excel chart background | set Excel chart background to PNG programmatically | Aspose.Cells load image bytes for chart fill
// Developer Intent: Use Aspose.Cells to programmatically replace the default chart background with a custom image file, such as a corporate logo.
// Use Cases: Branding reports by embedding a logo behind sales charts. | Creating a reusable template that watermarks all charts with a background image. | Allowing end‑users to select an image at runtime to personalize chart appearance. | Generating localized charts with region‑specific background graphics. | Automating compliance watermarks on financial charts.
// AI Prompts: Write C# code that loads a PNG from disk and sets it as the background of an Aspose.Cells chart, including error handling for missing files. | Explain the steps to use FillFormat.TextureFill.ImageData for chart area customization in Aspose.Cells. | Show how to change the FillType of a chart area to Texture and assign image bytes in a .NET workbook. | Provide a tutorial for adding a watermark image to every chart in an Excel file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This C# sample creates a workbook, adds a column chart, switches the chart area fill to a texture, reads a PNG/JPEG logo from a given path into a byte array, assigns the bytes to the chart’s TextureFill, and saves the file while gracefully handling a missing image.
class SetChartBackgroundImage
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Optional: add some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Set the chart area fill type to texture (image)
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            // Load the company logo image file into a byte array if it exists
            string logoPath = "company_logo.png"; // replace with actual logo file path
            if (File.Exists(logoPath))
            {
                byte[] logoBytes = File.ReadAllBytes(logoPath);
                // Assign the image data to the texture fill of the chart area
                chart.ChartArea.Area.FillFormat.TextureFill.ImageData = logoBytes;
            }
            else
            {
                Console.WriteLine($"Warning: Logo file '{logoPath}' not found. Chart background will use default texture.");
            }

            // Save the workbook with the chart background image applied
            string outputPath = "ChartWithLogoBackground.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
