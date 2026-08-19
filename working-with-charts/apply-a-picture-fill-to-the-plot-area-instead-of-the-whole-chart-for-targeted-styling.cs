// Title: Apply Picture Fill to a Chart Plot Area with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a column chart, and uses Aspose.Cells FillFormat to set a texture image as the plot‑area background. Demonstrates base64 image loading, FillPictureType.StackAndScale, scaling, and saving the file.
// Keywords: Aspose.Cells C# chart plot area fill | picture fill chart Aspose.Cells | FillFormat texture plot area | FillPictureType StackAndScale | chart background image .NET | base64 image Aspose.Cells | Excel chart styling C#
// Common Searches: how to set picture fill for chart plot area Aspose.Cells | Aspose.Cells fill chart background with image | C# set texture fill for Excel chart plot area | apply base64 image to chart plot area Aspose | FillFormat FillType Texture example
// Developer Intent: Style only the plot area of an Excel chart with a custom picture texture using Aspose.Cells for .NET.
// Use Cases: Add a column chart and give its plot area a custom textured background without affecting the chart frame. | Reuse the same image data to apply identical picture fills to multiple chart plot areas in a workbook. | Control the visual fit of the texture by adjusting the scaling factor or switching FillPictureType (e.g., Stretch, Tile).
// AI Prompts: Generate C# code that applies a JPEG file as a picture fill to the plot area of a line chart using Aspose.Cells, including options for scaling and tiling. | Show how to read an external PNG into a byte array and set it as the FillFormat.ImageData for a chart's plot area with FillPictureType.StackAndScale. | Provide an example that changes the picture fill type to Stretch for a pie chart's plot area and explains the visual impact.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a column chart, and uses Aspose.Cells FillFormat to set a texture image as the plot‑area background. Demonstrates base64 image loading, FillPictureType.StackAndScale, scaling, and saving the file.
class PlotAreaPictureFillDemo
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

        // ----- Apply picture fill to the plot area -----
        // Simple 1x1 white pixel image (base64 encoded)
        string base64Img = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
        byte[] imgData = Convert.FromBase64String(base64Img);

        // Set the fill type to texture (picture)
        chart.PlotArea.Area.FillFormat.FillType = FillType.Texture;
        // Assign the image data
        chart.PlotArea.Area.FillFormat.ImageData = imgData;
        // Choose how the picture is applied (stack and scale in this example)
        chart.PlotArea.Area.FillFormat.PictureFormatType = FillPictureType.StackAndScale;
        // Optional: set scaling factor (1.0 = original size)
        chart.PlotArea.Area.FillFormat.Scale = 1.0;

        // Save the workbook
        workbook.Save("PlotAreaPictureFill.xlsx");
    }
}
