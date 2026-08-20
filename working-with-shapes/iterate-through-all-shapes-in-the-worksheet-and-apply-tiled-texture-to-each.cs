// Title: Apply a Tiled Canvas Texture to Every Shape in Aspose.Cells (C#)
// Description: The sample creates a workbook, adds a few shapes, then iterates over the worksheet's Shapes collection. For each shape it sets FillType to Texture, selects the built‑in Canvas pattern, turns on tiling, configures TilePicOption (scale and offset), and finally saves the file as AllShapesTiledTexture.xlsx.
// Keywords: Aspose.Cells | C# shape fill | texture tiling | Canvas pattern | TilePicOption | iterate worksheet shapes | Excel workbook styling | repeat texture fill | programmatic Excel graphics | shape background texture
// Common Searches: Aspose.Cells C# apply texture fill to all shapes | Enable tiling for shape fills in Aspose.Cells | Set Canvas pattern as shape background using Aspose | Loop through worksheet shapes and configure TilePicOption | C# example for tiled texture in Excel with Aspose.Cells
// Developer Intent: Programmatically give every shape in a worksheet a repeated canvas texture.
// Use Cases: Standardize the visual style of diagram elements across a report by applying a common tiled background. | Create a template workbook where placeholder shapes automatically display a repeatable texture for branding purposes. | Process imported graphics and add a tiled texture before exporting the workbook to Excel for consistent presentation.
// AI Prompts: Generate C# code with Aspose.Cells that loops through all worksheet shapes and applies a tiled Canvas texture with custom scaling and offset. | Show how to switch the texture type to another built‑in pattern while keeping tiling enabled for each shape. | Explain how to modify TilePicOption properties to produce different tile sizes and positions across multiple shapes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, adds a few shapes, then iterates over the worksheet's Shapes collection. For each shape it sets FillType to Texture, selects the built‑in Canvas pattern, turns on tiling, configures TilePicOption (scale and offset), and finally saves the file as AllShapesTiledTexture.xlsx.
public class ApplyTiledTextureToAllShapes
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

        // Add sample shapes (optional, for demonstration)
        worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);
        worksheet.Shapes.AddOval(5, 0, 5, 0, 100, 100);
        worksheet.Shapes.AddTextBox(2, 0, 2, 0, 120, 200);

        // Apply tiled texture fill to each shape
        foreach (Shape shape in worksheet.Shapes)
        {
            shape.Fill.FillType = FillType.Texture;
            TextureFill textureFill = shape.Fill.TextureFill;

            // Use a built‑in texture type
            textureFill.Type = TextureType.Canvas;

            // Enable tiling so the texture repeats across the shape
            textureFill.IsTiling = true;

            // Define tile options (scaling and offset)
            TilePicOption tileOption = new TilePicOption
            {
                ScaleX = 0.5, // 50 % width scaling
                ScaleY = 0.5, // 50 % height scaling
                OffsetX = 5,  // horizontal offset
                OffsetY = 5   // vertical offset
            };
            textureFill.TilePicOption = tileOption;
        }

        // Save the workbook with the applied tiled textures
        workbook.Save("AllShapesTiledTexture.xlsx");
        Console.WriteLine("Workbook saved as AllShapesTiledTexture.xlsx");
    }
}
