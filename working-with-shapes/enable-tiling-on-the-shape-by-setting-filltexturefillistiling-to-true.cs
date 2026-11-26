// Title: Enable texture tiling on a rectangle shape with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set shape texture fill to tile | How to enable tiling for Fill.TextureFill in Aspose.Cells | C# example of repeating built‑in texture on a rectangle shape using Aspose.Cells | Use Fill.TextureFill.IsTiling property to tile texture in Aspose.Cells workbook | Create a tiled texture background for a shape and save as XLSX with Aspose.Cells
// Tags: Aspose.Cells shape texture tiling | Fill.TextureFill.IsTiling C# | Aspose.Cells rectangle shape texture fill | C# tiled texture fill example | Save workbook with tiled texture fill Xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureTilingDemo
{
    // The sample creates a workbook, adds a rectangle shape, sets its fill type to a built‑in texture (BlueTissuePaper), enables tiling by setting Fill.TextureFill.IsTiling to true, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, height (pixels), width (pixels)
                Shape rectangle = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 200, 300);

                // Set the fill type of the shape to Texture
                rectangle.Fill.FillType = FillType.Texture;

                // Obtain the TextureFill object
                TextureFill textureFill = rectangle.Fill.TextureFill;

                // Choose a built‑in texture type (you can also assign custom image data)
                textureFill.Type = TextureType.BlueTissuePaper;

                // Enable tiling of the texture
                textureFill.IsTiling = true;

                // Optional: demonstrate that the property is set
                Console.WriteLine("TextureFill.IsTiling = " + textureFill.IsTiling);

                // Save the workbook to a file
                string outputPath = Path.Combine(Environment.CurrentDirectory, "TextureTilingDemo.xlsx");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
