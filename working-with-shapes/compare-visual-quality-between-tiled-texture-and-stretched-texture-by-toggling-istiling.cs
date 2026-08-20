// Title: Toggle IsTiling to Compare Tiled vs Stretched Texture Fill on a Shape – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a rectangle shape, applies a PNG texture fill, then saves two files – one with IsTiling = true (repeating pattern) and another with IsTiling = false (stretched image) – to demonstrate the visual difference.
// Keywords: Aspose.Cells | C# | .NET | texture fill | IsTiling | tiled texture | stretched texture | shape fill | Excel shape texture | image data | PNG texture | workbook export
// Common Searches: Aspose.Cells how to enable texture tiling on a shape | IsTiling true vs false effect in Excel shape fill | C# code to compare tiled and stretched texture in Aspose.Cells | load PNG bytes into shape texture fill without a file | save separate workbooks for tiled and stretched shape textures
// Developer Intent: Show how changing the IsTiling property alters a shape's texture rendering by producing two workbooks – one with a repeated pattern and one with a stretched image.
// Use Cases: Design‑review workbook that visualizes both tiled and stretched pattern options. | Template generation where one version needs a repeatable background and another a full‑size image. | Automated documentation of branding guidelines that compares texture rendering modes.
// AI Prompts: Generate C# code to add a circular shape with a tiled texture fill using Aspose.Cells and save the workbook. | Explain the impact of the IsTiling property on texture rendering in Aspose.Cells and how to toggle it at runtime. | Provide a method to embed a PNG byte array directly into a shape's texture fill without writing the file to disk.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape, applies a PNG texture fill, then saves two files – one with IsTiling = true (repeating pattern) and another with IsTiling = false (stretched image) – to demonstrate the visual difference.
class CompareTextureTiling
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will demonstrate the texture fill
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 300, 200);
        shape.Fill.FillType = FillType.Texture;

        // Load texture image data (replace with a real file path if available)
        string imagePath = "texture.png";
        if (File.Exists(imagePath))
        {
            shape.Fill.TextureFill.ImageData = File.ReadAllBytes(imagePath);
        }
        else
        {
            // Fallback: a minimal 1x1 red pixel PNG encoded in base64
            shape.Fill.TextureFill.ImageData = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAF/AL+XKpZVQAAAABJRU5ErkJggg==");
        }

        // ---------- Tiled texture ----------
        // Enable tiling so the picture repeats across the shape
        shape.Fill.TextureFill.IsTiling = true;
        // Save workbook showing tiled texture
        workbook.Save("Texture_Tiled.xlsx");

        // ---------- Stretched texture ----------
        // Disable tiling; the picture will be stretched to fill the shape
        shape.Fill.TextureFill.IsTiling = false;
        // Save workbook showing stretched texture
        workbook.Save("Texture_Stretched.xlsx");
    }
}
