using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTests
{
    public class Program
    {
        public static void Main()
        {
            // Wrap the whole operation in a try‑catch to handle unexpected errors gracefully
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to demonstrate texture fill
                Shape rect = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);
                rect.Fill.FillType = FillType.Texture; // set fill type to texture

                // Configure the texture fill
                TextureFill textureFill = rect.Fill.TextureFill;
                textureFill.Type = TextureType.BlueTissuePaper; // built‑in texture
                textureFill.IsTiling = true;                     // enable tiling

                // Verify the property is set before saving
                if (!textureFill.IsTiling)
                    throw new InvalidOperationException("IsTiling should be true before saving.");

                // Save the workbook to a temporary file
                string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
                workbook.Save(tempPath);

                // Ensure the file was created before attempting to load it
                if (!File.Exists(tempPath))
                    throw new FileNotFoundException("Temporary workbook file was not found.", tempPath);

                // Load the workbook back
                Workbook loadedWorkbook = new Workbook(tempPath);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                Shape loadedRect = loadedSheet.Shapes[0]; // the rectangle we added

                // Retrieve the texture fill from the loaded shape
                TextureFill loadedTextureFill = loadedRect.Fill.TextureFill;

                // Verify that IsTiling persisted after reload
                if (!loadedTextureFill.IsTiling)
                    throw new InvalidOperationException("IsTiling should remain true after reloading the workbook.");

                Console.WriteLine("Texture fill tiling property persisted successfully.");
                
                // Clean up the temporary file
                File.Delete(tempPath);
            }
            catch (Exception ex)
            {
                // Output any errors that occurred during execution
                Console.WriteLine($"Error: {ex.GetType().Name} - {ex.Message}");
            }
        }
    }
}