// Title: Create a C# unit test with Aspose.Cells that verifies texture tiling is enabled on a rectangle shape when IsTiling is true
// AI Prompts: Generate a C# NUnit/MSTest method that adds a rectangle shape to a worksheet, sets shape.Fill.TextureFill.IsTiling = true, and asserts the property equals true. | Write code that saves the workbook to a MemoryStream after enabling texture tiling to confirm the file can be written without errors.
// Common Searches: how to unit test texture fill tiling in Aspose.Cells C# | Aspose.Cells verify Fill.TextureFill.IsTiling property in a test | C# example for asserting rectangle shape texture tiling with Aspose.Cells
// Tags: Aspose.Cells shape texture tiling unit test | C# Fill.TextureFill.IsTiling verification | Aspose.Cells save workbook to MemoryStream | rectangle shape texture fill Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTests
{
    // The example creates a Workbook, adds a rectangle shape, enables texture tiling via Fill.TextureFill.IsTiling = true, asserts the property is true, saves the workbook to a MemoryStream, and confirms the operation succeeds.
    public class TextureTilingTests
    {
        public void Run()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Get the first worksheet
                var sheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                var shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 100);

                // Enable texture tiling using the updated Fill API
                shape.Fill.TextureFill.IsTiling = true;

                // Verify that the IsTiling property is true
                if (!shape.Fill.TextureFill.IsTiling)
                {
                    throw new InvalidOperationException("Texture tiling should be enabled when IsTiling is set to true.");
                }

                // Save to a memory stream to ensure the workbook can be saved without errors
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                }

                Console.WriteLine("Texture tiling test passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during the texture tiling test: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            var test = new TextureTilingTests();
            test.Run();
        }
    }
}
