// Title: Copy a Shape and Apply a New Texture Image with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a rectangle shape, optionally set a PNG texture, duplicate the shape to a different location, and assign a distinct JPG texture (with optional tiling) to the copy before saving the Excel file.
// Keywords: Aspose.Cells | C# shape copy | texture fill | Excel shape duplicate | image fill Aspose.Cells | AddCopy method | FillType.Texture | worksheet shapes | programmatic shape duplication | Excel file generation
// Common Searches: Aspose.Cells copy shape and change fill | C# duplicate Excel shape with different image | How to set texture fill for a shape in Aspose.Cells | AddCopy shape Aspose.Cells example | Replace shape texture after copying
// Developer Intent: Programmatically duplicate an existing worksheet shape and give the duplicate a distinct texture image.
// Use Cases: Create a product catalog where each item icon is a copied shape with its own image. | Design a dashboard with repeated shape placeholders that need individual background pictures. | Generate a patterned worksheet by tiling different textures on copied shapes. | Automate branding by copying a logo shape and swapping its fill for region‑specific graphics.
// AI Prompts: Generate C# code using Aspose.Cells to copy a rectangle shape and assign a new JPEG texture to the copy, handling missing files gracefully. | Show an example that duplicates a shape, sets FillType.Texture, and enables tiling for the copied shape in a .xlsx workbook. | Explain how to use Worksheet.Shapes.AddCopy to replicate a shape and then change its ImageData property.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a rectangle shape, optionally set a PNG texture, duplicate the shape to a different location, and assign a distinct JPG texture (with optional tiling) to the copy before saving the Excel file.
class DuplicateShapeWithTexture
{
    static void Main()
    {
        try
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add an original rectangle shape
            Shape originalShape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 130, 130);

            // (Optional) Apply a texture to the original shape
            originalShape.Fill.FillType = FillType.Texture;
            string originalTexturePath = "originalTexture.png";
            if (File.Exists(originalTexturePath))
            {
                originalShape.Fill.TextureFill.ImageData = File.ReadAllBytes(originalTexturePath);
            }

            // Duplicate the shape to a new position
            Shape copiedShape = worksheet.Shapes.AddCopy(originalShape, 7, 0, 7, 0);

            // Assign a different texture image to the copied shape (if the file exists)
            copiedShape.Fill.FillType = FillType.Texture;
            string newTexturePath = "newTexture.jpg";
            if (File.Exists(newTexturePath))
            {
                copiedShape.Fill.TextureFill.ImageData = File.ReadAllBytes(newTexturePath);
                copiedShape.Fill.TextureFill.IsTiling = true; // optional tiling
            }

            // Save the workbook
            workbook.Save("DuplicatedShapeWithTexture.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
