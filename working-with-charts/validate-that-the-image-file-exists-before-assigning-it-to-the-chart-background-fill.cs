// Title: C# – Validate Image File Before Setting Chart Background Texture in Aspose.Cells
// Description: The sample builds a workbook, inserts sample data, creates a column chart, checks that a PNG file exists, loads it into a byte array, applies the image as a texture fill to the chart area, and saves the workbook with the custom background.
// Keywords: Aspose.Cells chart background image | C# validate image file existence | chart area texture fill Aspose.Cells | set chart background from bytes | Excel chart custom background .NET | load PNG into byte array Aspose.Cells | file existence check before chart fill | Aspose.Cells .NET example
// Common Searches: Aspose.Cells verify image before chart background | C# set chart area fill with image Aspose.Cells | how to use texture fill for chart background in Aspose.Cells | check file existence when applying chart background | load PNG as chart background Aspose.Cells .NET
// Developer Intent: Ensure the image file is present and readable before using it as a chart background texture.
// Use Cases: Generate Excel reports where charts display custom background images only when the images are available. | Create a helper method that validates an image path and applies the texture to multiple charts in the same workbook. | Select different chart backgrounds at runtime based on which image files exist on the server.
// AI Prompts: Write C# code using Aspose.Cells that checks for a PNG file, reads it into a byte array, and applies it as the chart area texture fill. | Create a reusable function that takes a Chart object and an image path, validates the file, sets FillType to Texture, assigns ImageData, and returns a success or error message. | Show error‑handling patterns for missing image files when setting a chart background in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The sample builds a workbook, inserts sample data, creates a column chart, checks that a PNG file exists, loads it into a byte array, applies the image as a texture fill to the chart area, and saves the workbook with the custom background.
class SetChartBackgroundWithValidation
{
    static void Main()
    {
        // Path to the image file that will be used as chart background
        string imagePath = "chartBackground.png";

        // Validate that the image file exists before proceeding
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Image file not found: {imagePath}");
            return;
        }

        // Read the image file into a byte array
        byte[] imageData = File.ReadAllBytes(imagePath);

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);

        // Set the chart area fill type to texture and assign the image data
        chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;
        chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;

        // Save the workbook with the chart background image
        workbook.Save("ChartWithBackground.xlsx");
        Console.WriteLine("Workbook saved successfully with chart background image.");
    }
}
