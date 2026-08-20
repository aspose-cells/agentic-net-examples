// Title: Tile a picture as texture inside a rectangle shape while preserving aspect ratio with Aspose.Cells (C#)
// Description: Demonstrates how to add a rectangle shape to a workbook, load a PNG image as a texture, enable tiling, keep the image's original aspect ratio using FillPictureType.StackAndScale, optionally scale each tile, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells texture fill | C# shape tiling | preserve aspect ratio fill | FillPictureType.StackAndScale | TilePicOption scaling | rectangle shape texture | Aspose.Cells example | image tiling spreadsheet
// Common Searches: Aspose.Cells tile image inside shape C# | preserve aspect ratio texture fill Aspose.Cells | how to scale tiled picture in Aspose.Cells | FillPictureType options Aspose.Cells .NET | add rectangle shape with texture fill Aspose.Cells
// Developer Intent: Apply a tiled image texture to a rectangle shape in a spreadsheet while maintaining the image's original proportions.
// Use Cases: Create a repeating logo background for a report header without distortion. | Design patterned cells or banners where the texture repeats uniformly. | Generate custom worksheet graphics that require scaled tiles while keeping the source image's aspect ratio.
// AI Prompts: Modify the sample to use a 75% tile scale and turn off tiling. | Show how to stretch the texture instead of preserving aspect ratio using a different FillPictureType. | Load the texture image from a MemoryStream rather than a file path in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape to a workbook, load a PNG image as a texture, enable tiling, keep the image's original aspect ratio using FillPictureType.StackAndScale, optionally scale each tile, and save the file with Aspose.Cells for .NET.
public class TextureFillTilingDemo
{
    public static void Main(string[] args)
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
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape where the texture will be applied
        // Parameters: upper left row, upper left column, top, left, width, height (in points)
        Shape rectangle = worksheet.Shapes.AddRectangle(2, 2, 2, 2, 300, 200);

        // Set the fill type of the shape to texture
        rectangle.Fill.FillType = FillType.Texture;

        // Load the picture file that will be used as the texture
        string texturePath = "texture.png";
        if (File.Exists(texturePath))
        {
            byte[] imageData = File.ReadAllBytes(texturePath);
            rectangle.Fill.TextureFill.ImageData = imageData;
        }
        else
        {
            Console.WriteLine($"Warning: Texture file '{texturePath}' not found. Skipping texture fill.");
        }

        // Enable tiling so the picture repeats to fill the shape
        rectangle.Fill.TextureFill.IsTiling = true;

        // Preserve the aspect ratio of the picture while tiling.
        // Using FillPictureType.StackAndScale keeps the original aspect ratio.
        rectangle.Fill.TextureFill.PictureFormatType = FillPictureType.StackAndScale;

        // Optionally adjust the scale of each tile (same value for X and Y to keep aspect ratio)
        // Here we set each tile to 50% of its original size.
        TilePicOption tileOptions = new TilePicOption
        {
            ScaleX = 0.5,   // 50% horizontal scale
            ScaleY = 0.5    // 50% vertical scale
        };
        rectangle.Fill.TextureFill.TilePicOption = tileOptions;

        // Save the workbook with the textured rectangle
        string outputPath = "TextureFillTilingDemo.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
