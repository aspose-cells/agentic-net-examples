// Title: C# unit test to confirm texture fill tiling (IsTiling) persists in an Aspose.Cells workbook
// Common Searches: how to assert texture fill tiling is saved in Aspose.Cells C# unit test | Aspose.Cells verify Shape Fill.IsTiling property after workbook save | C# example testing texture fill tiling persistence with memory stream in Aspose.Cells | unit testing Shape Fill.IsTiling flag in an Excel file using Aspose.Cells | check texture fill tiling flag remains true after reloading workbook Aspose.Cells
// Tags: Aspose.Cells texture fill persistence test | C# unit test Shape Fill.IsTiling | Excel workbook save reload texture properties | MemoryStream based Aspose.Cells unit testing | FillType.Texture shape verification in C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureTilingTest
{
    // The example creates a Workbook, adds a rectangle shape, sets its FillType to Texture, selects a built‑in texture, enables IsTiling, saves the workbook to a MemoryStream, reloads it, and asserts that the loaded shape's TextureFill.IsTiling remains true, throwing an exception if the flag is not persisted.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                TestTextureTilingApplied();
                Console.WriteLine("All tests completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
                Environment.Exit(1);
            }
        }

        static void TestTextureTilingApplied()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape (parameters: upper left row, upper left column, upper left offset X,
            // upper left offset Y, width, height)
            Shape rectangle = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 200, 100);

            // Configure the shape to use a texture fill
            rectangle.Fill.FillType = FillType.Texture;
            TextureFill textureFill = rectangle.Fill.TextureFill;

            // Use a built‑in texture type to avoid external image files
            textureFill.Type = TextureType.BlueTissuePaper;

            // Enable tiling
            textureFill.IsTiling = true;

            // Save the workbook to a memory stream (no file I/O required)
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0; // Reset stream position for reading

                // Reload the workbook from the stream
                Workbook loadedWorkbook = new Workbook(ms);
                Shape loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];

                // Retrieve the IsTiling flag from the loaded shape
                bool isTilingPersisted = loadedShape.Fill.TextureFill.IsTiling;

                // Verify the flag is true; throw if the expectation is not met
                if (!isTilingPersisted)
                {
                    throw new InvalidOperationException(
                        "TextureFill.IsTiling was expected to be true after reload, but it was false.");
                }
            }

            Console.WriteLine("TestTextureTilingApplied passed: IsTiling is true and persisted correctly.");
        }
    }
}
