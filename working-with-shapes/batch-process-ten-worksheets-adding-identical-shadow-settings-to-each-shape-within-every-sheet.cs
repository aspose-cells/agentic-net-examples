// Title: Apply the same shadow effect to every shape in the first ten worksheets of an Excel file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to set a uniform shadow on all shapes across multiple worksheets | batch apply shadow effect to Excel shapes with Aspose.Cells .NET | iterate through first ten worksheets and add shadow to each shape using Aspose.Cells | set preset shadow type and color for shapes in a workbook with Aspose.Cells C# | apply identical shadow parameters to shapes in several Excel sheets programmatically
// Tags: set shape shadow effect Aspose.Cells | batch apply shadow to Excel shapes | preset shadow type configuration C# | process multiple worksheets shapes | apply uniform shadow color to shapes | shadow distance blur transparency Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace BatchShadowDemo
{
    // The example loads or creates a workbook, ensures at least ten worksheets, adds a rectangle shape where needed, defines a preset shadow (type, distance, blur, transparency, angle, gray color), then loops through the first ten worksheets applying the same shadow settings to every shape before saving the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (adjust as needed)
            string inputPath = "input.xlsx";
            string outputPath = "output_with_shadows.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    workbook = new Workbook();
                    Console.WriteLine("Created a new workbook.");
                }

                // Ensure there are at least 10 worksheets
                const int requiredSheets = 10;
                while (workbook.Worksheets.Count < requiredSheets)
                {
                    workbook.Worksheets.Add();
                }

                // Optional: Add a sample shape to each worksheet if it has none
                for (int i = 0; i < requiredSheets; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    if (sheet.Shapes.Count == 0)
                    {
                        // Add a simple rectangle shape for demonstration
                        // Parameters: upper left row, upper left column, top offset, left offset, height, width
                        sheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 150);
                        Console.WriteLine($"Added a rectangle to worksheet \"{sheet.Name}\" (index {i}).");
                    }
                }

                // Define the identical shadow settings to apply
                // These settings will be applied to every shape in every of the first ten worksheets
                PresetShadowType presetType = PresetShadowType.OffsetDiagonalBottomRight;
                double shadowDistance = 10.0;   // points
                double shadowBlur = 5.0;        // points
                double shadowTransparency = 0.4; // 0 (opaque) to 1 (clear)
                double shadowAngle = 45.0;      // degrees

                // Create a semi‑transparent gray color for the shadow
                // ARGB: Alpha (128) = 50% opacity, Gray (128,128,128)
                CellsColor shadowColor = workbook.CreateCellsColor();
                shadowColor.Argb = (128 << 24) | (128 << 16) | (128 << 8) | 128;

                // Process the first ten worksheets
                for (int sheetIndex = 0; sheetIndex < requiredSheets; sheetIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[sheetIndex];
                    ShapeCollection shapes = sheet.Shapes;

                    Console.WriteLine($"Processing worksheet \"{sheet.Name}\" (index {sheetIndex}) with {shapes.Count} shape(s).");

                    for (int shapeIdx = 0; shapeIdx < shapes.Count; shapeIdx++)
                    {
                        Shape shape = shapes[shapeIdx];

                        // Access the ShadowEffect of the shape
                        ShadowEffect shadow = shape.ShadowEffect;

                        // Apply the predefined settings
                        shadow.PresetType = presetType;
                        shadow.Distance = shadowDistance;
                        shadow.Blur = shadowBlur;
                        shadow.Transparency = shadowTransparency;
                        shadow.Angle = shadowAngle;
                        shadow.Color = shadowColor;

                        Console.WriteLine($"  Applied shadow to shape \"{shape.Name}\" (type {shape.Type}).");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with shadow effects to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during processing:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
