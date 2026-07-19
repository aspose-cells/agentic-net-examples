// Title: C# – Validate Image File Before Applying Chart Background in Aspose.Cells
// Description: Creates a workbook, adds a column chart, and sets the chart area fill to a texture using a local PNG file. The code first checks File.Exists, reads the image into a byte array, and applies it only when the file is present, then saves the workbook.
// Keywords: Aspose.Cells chart background image | C# validate image file exists | texture fill chart Aspose.Cells | chart area fill from file | .NET Excel chart image | file existence check Aspose
// Common Searches: how to check image file before using it as chart background Aspose.Cells | set chart background texture from local image C# | Aspose.Cells verify file exists before chart fill | apply PNG as chart area fill Aspose.Cells .NET | error handling for missing chart background image
// Developer Intent: Ensure the specified image file is present before assigning it as a texture fill for a chart background.
// Use Cases: Load a PNG, confirm its existence, and use it as the background of a newly created column chart. | Skip background image assignment when the file is missing while still generating the workbook. | Reuse the validation routine to apply different images to multiple charts in the same workbook.
// AI Prompts: Write C# code that adds a chart with Aspose.Cells and sets its background image only after confirming the file exists, including proper error handling. | Create a reusable method that takes a Chart object and an image path, validates the file, and applies the image as a texture fill. | Explain how to fall back to a solid color fill when the background image file cannot be found while using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a column chart, and sets the chart area fill to a texture using a local PNG file. The code first checks File.Exists, reads the image into a byte array, and applies it only when the file is present, then saves the workbook.
class SetChartBackgroundImage
{
    static void Main()
    {
        // Path to the image file that will be used as chart background
        string imagePath = "chart_background.png";

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

        // Optional: add some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B3", true);

        // Configure the chart background fill to use a texture and assign the image data
        chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;
        chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;

        // Save the workbook with the chart that now has a background image
        workbook.Save("ChartWithBackgroundImage.xlsx");
        Console.WriteLine("Workbook saved with chart background image.");
    }
}
