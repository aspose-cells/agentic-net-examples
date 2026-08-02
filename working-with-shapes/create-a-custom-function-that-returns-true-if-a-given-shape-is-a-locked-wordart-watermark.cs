// Title: C# Extension Method to Detect Locked WordArt Watermarks in Excel using Aspose.Cells
// Description: Provides a ShapeExtensions class with an IsLockedWordArtWatermark extension method that returns true when a Shape's IsWordArt and IsLocked properties are both set. The example shows how to load a workbook, iterate worksheet shapes, and identify locked WordArt watermarks with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# extension method | Excel Shape | WordArt watermark | locked shape detection | Shape.IsWordArt | Shape.IsLocked | .NET workbook processing | watermark identification | Excel automation
// Common Searches: Aspose.Cells detect locked WordArt watermark | C# extension method shape.IsWordArt IsLocked | how to find watermarks in Excel with Aspose.Cells | identify WordArt watermarks in a workbook | check if Excel shape is locked WordArt using .NET | filter watermarks from Excel sheets programmatically
// Developer Intent: Determine whether a given Shape object represents a locked WordArt watermark in an Excel worksheet.
// Use Cases: Scan all shapes in a worksheet and list those that are locked WordArt watermarks before publishing the file. | Programmatically remove or hide locked WordArt watermarks by deleting shapes that match the IsLockedWordArtWatermark criteria. | Validate incoming Excel files to ensure they do not contain prohibited locked WordArt watermarks for compliance checks.
// AI Prompts: Generate a C# extension method for Aspose.Cells Shape that returns true when the shape is both WordArt and locked. | Show sample code that loads an Excel workbook, iterates over worksheet shapes, and uses IsLockedWordArtWatermark to filter watermarks. | Explain how to combine the IsLockedWordArtWatermark method with shape removal logic to clean a workbook of watermarks.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Provides a ShapeExtensions class with an IsLockedWordArtWatermark extension method that returns true when a Shape's IsWordArt and IsLocked properties are both set. The example shows how to load a workbook, iterate worksheet shapes, and identify locked WordArt watermarks with Aspose.Cells for .NET.
    public static class ShapeExtensions
    {
        // Determines whether the specified shape is a locked WordArt watermark.
        // A typical watermark in Excel is implemented as a WordArt shape that is locked.
        public static bool IsLockedWordArtWatermark(this Shape shape)
        {
            // Guard against null references.
            if (shape == null)
                return false;

            // Check both WordArt flag and locked flag.
            return shape.IsWordArt && shape.IsLocked;
        }
    }

    internal class Program
    {
        // Entry point required for compilation.
        private static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Aspose.Cells ShapeExtensions loaded successfully.");

                // Example usage (optional):
                // Load a workbook if a sample file exists.
                string samplePath = "Sample.xlsx";
                if (File.Exists(samplePath))
                {
                    var workbook = new Workbook(samplePath);
                    var worksheet = workbook.Worksheets[0];
                    var shapes = worksheet.Shapes;

                    foreach (Shape shape in shapes)
                    {
                        bool isWatermark = shape.IsLockedWordArtWatermark();
                        Console.WriteLine($"Shape '{shape.Name}' is locked WordArt watermark: {isWatermark}");
                    }
                }
                else
                {
                    Console.WriteLine($"Sample file '{samplePath}' not found. Skipping workbook processing.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
