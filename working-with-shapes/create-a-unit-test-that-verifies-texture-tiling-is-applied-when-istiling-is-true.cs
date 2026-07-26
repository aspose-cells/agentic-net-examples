// Title: C# Unit Test: Verify Aspose.Cells Shape TextureFill.IsTiling Persists After Save/Load
// Description: Shows how to create an Aspose.Cells workbook, add a rectangle shape, set Fill.FillType to Texture and enable TextureFill.IsTiling, save the file, reload it, and assert that the IsTiling flag remains true. Includes temporary‑file cleanup.
// Keywords: Aspose.Cells | C# | TextureFill | IsTiling | unit test | shape fill | texture tiling | MSTest | NUnit | xUnit | Excel automation
// Common Searches: Aspose.Cells unit test texture tiling | C# verify TextureFill.IsTiling after workbook save | how to test shape fill properties in Aspose.Cells | persist texture fill tiling Aspose.Cells | unit testing Aspose.Cells shape fill
// Developer Intent: Ensure the TextureFill.IsTiling property of a shape stays true after the workbook is saved and reopened.
// Use Cases: Regression test for Excel reports that use tiled texture backgrounds | CI pipeline check that shape fill settings are not lost during serialization | Quality assurance of custom templates relying on texture fill tiling
// AI Prompts: Generate an MSTest method that creates a workbook, adds a rectangle shape with TextureFill.IsTiling = true, saves and reloads the file, then asserts the property is true. | Write a NUnit test case for Aspose.Cells verifying that a shape's TextureFill.IsTiling flag persists after persisting the workbook. | Provide an xUnit test example that checks texture fill tiling is applied to a shape and remains after workbook serialization.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTests
{
    // Shows how to create an Aspose.Cells workbook, add a rectangle shape, set Fill.FillType to Texture and enable TextureFill.IsTiling, save the file, reload it, and assert that the IsTiling flag remains true. Includes temporary‑file cleanup.
    public class TextureFillTilingDemo
    {
        public static void Main()
        {
            string tempFile = null;

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to demonstrate texture fill
                Shape rectangle = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);
                rectangle.Fill.FillType = FillType.Texture;
                rectangle.Fill.TextureFill.IsTiling = true; // Set tiling to true

                // Save the workbook to a temporary file
                tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                workbook.Save(tempFile);

                // Ensure the file exists before loading
                if (!File.Exists(tempFile))
                    throw new FileNotFoundException("Temporary workbook file not found.", tempFile);

                // Load the workbook back from the file
                Workbook loadedWorkbook = new Workbook(tempFile);
                Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
                Shape loadedRectangle = loadedWorksheet.Shapes[0];

                // Verify that the IsTiling property is still true after reload
                bool isTiling = loadedRectangle.Fill.TextureFill.IsTiling;
                Console.WriteLine(isTiling
                    ? "Success: TextureFill.IsTiling is true after reloading the workbook."
                    : "Failure: TextureFill.IsTiling is false after reloading the workbook.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary file if it was created
                if (!string.IsNullOrEmpty(tempFile) && File.Exists(tempFile))
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch
                    {
                        // Suppress any exceptions during cleanup
                    }
                }
            }
        }
    }
}
