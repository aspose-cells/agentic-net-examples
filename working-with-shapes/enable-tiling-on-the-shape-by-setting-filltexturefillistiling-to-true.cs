// Title: Enable Texture Tiling on a Shape with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds a rectangle shape, sets its fill type to Texture, selects a built‑in texture (BlueTissuePaper), enables tiling by setting TextureFill.IsTiling to true, and saves the file as EnableTilingDemo.xlsx.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# shape texture fill | TextureFill.IsTiling | shape tiling | AddRectangle | FillType.Texture | TextureType.BlueTissuePaper | worksheet shape fill | code example | EnableTilingDemo
// Common Searches: Aspose.Cells enable texture tiling on shape C# | How to set IsTiling for a shape fill in Aspose.Cells | C# Aspose.Cells texture fill repeat pattern | AddRectangle with tiled texture Aspose.Cells | TextureFill.IsTiling property usage Aspose.Cells .NET
// Developer Intent: Turn on texture tiling for a worksheet shape using Aspose.Cells.
// Use Cases: Create a patterned background by repeating a texture across a shape. | Apply a tiled logo or watermark texture to a shape for branding. | Design mock‑ups that require fabric or paper textures tiled inside shapes.
// AI Prompts: Write C# code with Aspose.Cells that adds a rectangle shape and applies a tiled built‑in texture fill. | Show how to toggle TextureFill.IsTiling on and off for a shape in an Aspose.Cells workbook. | Provide an Aspose.Cells example that uses a custom image as a tiled texture fill for a shape and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, adds a rectangle shape, sets its fill type to Texture, selects a built‑in texture (BlueTissuePaper), enables tiling by setting TextureFill.IsTiling to true, and saves the file as EnableTilingDemo.xlsx.
class EnableTilingDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);

        // Set the shape's fill type to texture so we can access TextureFill
        shape.Fill.FillType = FillType.Texture;

        // Get the TextureFill object associated with the shape
        TextureFill textureFill = shape.Fill.TextureFill;

        // (Optional) Choose a built‑in texture type
        textureFill.Type = TextureType.BlueTissuePaper;

        // Enable tiling of the texture
        textureFill.IsTiling = true;

        // Save the workbook to a file
        workbook.Save("EnableTilingDemo.xlsx");
    }
}
