// Title: Aspose.Cells .NET unit test for verifying texture fill tiling (IsTiling) on a shape
// Description: Creates a workbook, adds a rectangle shape, sets FillType to Texture, enables IsTiling, saves to a temporary file, reloads the workbook, and asserts that the IsTiling flag remains true, confirming that texture tiling persists through serialization.
// Keywords: Aspose.Cells | .NET | texture fill | IsTiling | unit test | shape fill | rectangle shape | fill persistence | MSTest | xUnit | NUnit
// Common Searches: Aspose.Cells texture fill tiling unit test | How to test IsTiling property in Aspose.Cells | Verify texture fill persists after saving workbook | Shape fill tiling example Aspose.Cells .NET
// Developer Intent: Write an automated test that confirms the IsTiling property on a shape’s texture fill is set and retained after workbook save/load.
// Use Cases: Validate texture tiling for shapes in generated Excel reports. | Ensure workbook serialization keeps texture fill settings for downstream processing. | Add regression coverage for shape fill properties in CI pipelines.
// AI Prompts: Generate an MSTest method that creates a workbook, adds a rectangle shape, sets FillType.Texture, enables IsTiling, saves and reloads the file, then asserts IsTiling is true. | Write an xUnit test for Aspose.Cells that checks texture fill tiling persists after a round‑trip save. | Provide a NUnit example that verifies the TextureFill.IsTiling flag on a shape before and after workbook serialization.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTests
{
    // Creates a workbook, adds a rectangle shape, sets FillType to Texture, enables IsTiling, saves to a temporary file, reloads the workbook, and asserts that the IsTiling flag remains true, confirming that texture tiling persists through serialization.
    public class Program
    {
        public static void Main()
        {
            // Run the texture fill tiling demonstration
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to demonstrate texture fill
                Shape rectangle = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);

                // Configure the shape to use texture fill
                rectangle.Fill.FillType = FillType.Texture;
                // Use a built‑in texture type (any type works for the demo)
                rectangle.Fill.TextureFill.Type = TextureType.BlueTissuePaper;

                // Enable tiling
                rectangle.Fill.TextureFill.IsTiling = true;

                // Verify that the property is set to true
                if (!rectangle.Fill.TextureFill.IsTiling)
                {
                    throw new InvalidOperationException("IsTiling should be true after setting it.");
                }

                // Save the workbook to a temporary file
                string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                workbook.Save(tempFile);

                // Ensure the file was created before attempting to load
                if (!File.Exists(tempFile))
                {
                    throw new FileNotFoundException("Temporary workbook file was not created.", tempFile);
                }

                // Reload the workbook and retrieve the same shape
                Workbook loadedWorkbook = new Workbook(tempFile);
                Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
                Shape loadedRectangle = loadedWorksheet.Shapes[0];

                // Verify that the IsTiling property persisted after saving and loading
                if (!loadedRectangle.Fill.TextureFill.IsTiling)
                {
                    throw new InvalidOperationException("IsTiling should remain true after saving and loading the workbook.");
                }

                Console.WriteLine("Texture fill tiling test passed successfully.");
                
                // Clean up the temporary file
                File.Delete(tempFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
