// Title: Compare tiled versus stretched texture fill on rectangle shapes with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set IsTiling true for texture fill on a shape | C# example of applying tiled texture to a rectangle shape in Excel using Aspose.Cells | How to stretch a texture fill on an Excel shape with Aspose.Cells .NET | Create labeled rectangle shapes with texture fill and compare tiling in Aspose.Cells | Aspose.Cells texture fill scaling and tiling properties tutorial
// Tags: texture fill tiling Aspose.Cells | rectangle shape texture fill C# | IsTiling property Aspose.Cells | label shapes with text boxes Excel C# | texture fill scaling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, adds two rectangle shapes, applies the same PNG texture as a fill, sets IsTiling to true for the first shape (tiled) and false for the second (stretched), adds text‑box labels identifying each shape, and saves the file as 'TextureTilingComparison.xlsx'.
class Program
{
    // Base64 encoded 8x8 PNG (a simple checker pattern) used as texture.
    private const string Base64Texture =
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAYAAADED76LAAAAFklEQVQYV2NkYGD4z0AEYBxVSFIAABcAAf+XKcUAAAAASUVORK5CYII=";

    static void Main()
    {
        try
        {
            // Decode the texture image.
            byte[] textureData = Convert.FromBase64String(Base64Texture);

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Shape 1 – Tiled texture (IsTiling = true)
            // -------------------------------------------------
            Shape tiledShape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 300, 200);
            tiledShape.Fill.FillType = FillType.Texture;
            TextureFill tiledFill = tiledShape.Fill.TextureFill;
            tiledFill.ImageData = textureData;
            tiledFill.IsTiling = true;               // Tile the picture.
            tiledFill.Transparency = 0.0;            // Fully opaque.
            tiledFill.Scale = 1.0;                   // No additional scaling.
            tiledShape.AlternativeText = "Tiled texture";

            // -------------------------------------------------
            // Shape 2 – Stretched texture (IsTiling = false)
            // -------------------------------------------------
            Shape stretchedShape = sheet.Shapes.AddRectangle(5, 0, 1, 0, 300, 200);
            stretchedShape.Fill.FillType = FillType.Texture;
            TextureFill stretchedFill = stretchedShape.Fill.TextureFill;
            stretchedFill.ImageData = textureData;
            stretchedFill.IsTiling = false;          // Stretch the picture.
            stretchedFill.Transparency = 0.0;
            stretchedFill.Scale = 1.0;
            stretchedShape.AlternativeText = "Stretched texture";

            // Add labels so the user can see which shape is which.
            AddLabel(sheet, "Tiled (IsTiling = true)", 1, 0, 0, 0);
            AddLabel(sheet, "Stretched (IsTiling = false)", 5, 0, 0, 0);

            // Save the workbook.
            string outFile = "TextureTilingComparison.xlsx";
            workbook.Save(outFile);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outFile)}'.");

            // Simple verification output.
            Console.WriteLine($"Tiled shape IsTiling: {tiledFill.IsTiling}");
            Console.WriteLine($"Stretched shape IsTiling: {stretchedFill.IsTiling}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("An error occurred: " + ex.Message);
            Console.Error.WriteLine(ex.StackTrace);
        }
    }

    private static void AddLabel(Worksheet sheet, string text, int row, int col, int offsetX, int offsetY)
    {
        // Position the label a few rows above the target shape.
        Shape label = sheet.Shapes.AddTextBox(row, col, offsetX, offsetY, 300, 30);
        label.Text = text;
        label.Font.Size = 12;
        label.Font.IsBold = true;
        label.Fill.FillType = FillType.Solid;
        // Use ForeColor for solid fill (SolidColor is not a valid property).
        label.Fill.ForeColor = System.Drawing.Color.LightYellow;
        label.Line.Weight = 0.5;
    }
}
