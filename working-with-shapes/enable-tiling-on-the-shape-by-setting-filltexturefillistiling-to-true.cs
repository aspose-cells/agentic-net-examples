// Title: Enable texture tiling on a shape with Fill.TextureFill.IsTiling in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add a rectangle shape to a worksheet, set its fill to a built‑in texture (BlueTissuePaper), turn on tiling via the IsTiling property, and save the workbook as ShapeTextureTiling.xlsx.
// Keywords: Aspose.Cells | C# shape texture tiling | Fill.TextureFill.IsTiling | texture fill example | rectangle shape texture | BlueTissuePaper texture | .NET spreadsheet graphics | Aspose.Cells FillType.Texture | shape fill tiling
// Common Searches: Aspose.Cells enable texture tiling on shape | C# Fill.TextureFill.IsTiling usage | how to repeat texture fill in Aspose.Cells | set IsTiling property for shape fill .NET | texture fill tiling example Aspose.Cells
// Developer Intent: Activate tiling for a shape's texture fill.
// Use Cases: Create a patterned background by repeating a texture across a shape. | Maintain consistent texture appearance when resizing shapes in generated reports. | Apply the same tiled texture to multiple shapes for a uniform design.
// AI Prompts: Show a C# snippet that enables texture tiling on a rectangle shape using Aspose.Cells. | Explain how to toggle the IsTiling property for different built‑in textures in Aspose.Cells. | Provide step‑by‑step instructions to apply a tiled texture fill to shapes in a .NET workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape to a worksheet, set its fill to a built‑in texture (BlueTissuePaper), turn on tiling via the IsTiling property, and save the workbook as ShapeTextureTiling.xlsx.
class EnableTextureTilingDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);

        // Set the fill type of the shape to texture
        shape.Fill.FillType = FillType.Texture;

        // Get the TextureFill object associated with the shape
        TextureFill textureFill = shape.Fill.TextureFill;

        // (Optional) Choose a built‑in texture type
        textureFill.Type = TextureType.BlueTissuePaper;

        // Enable tiling for the texture fill
        textureFill.IsTiling = true;

        // Save the workbook to a file
        workbook.Save("ShapeTextureTiling.xlsx");
    }
}
