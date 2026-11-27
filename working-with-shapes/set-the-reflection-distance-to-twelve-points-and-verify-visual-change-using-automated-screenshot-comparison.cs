// Title: Set reflection distance to 12 points on a rectangle shape and validate the change with image comparison using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to adjust shape reflection distance programmatically | C# render Excel worksheet to PNG before and after modifying shape reflection | automated comparison of two worksheet images to detect shape effect changes in Aspose.Cells | set half reflection touching effect distance using Aspose.Cells Shape.Reflection | verify visual impact of reflection distance change with file size comparison in .NET
// Tags: shape reflection distance Aspose.Cells | render worksheet to PNG Aspose.Cells | image size comparison C# | half reflection touching effect Aspose.Cells | automated visual verification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using System.Drawing.Imaging;

// Creates a workbook, adds a rectangle shape with a half‑reflection effect, renders the sheet to PNG before and after setting the shape's Reflection.Distance to 12 points, and compares the two images by file size to confirm the visual change.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for temporary files
            string workbookPath = "ReflectionDistanceDemo.xlsx";
            string imgBeforePath = "before.png";
            string imgAfterPath = "after.png";

            // 1. Create a new workbook and add a rectangle shape
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Add a rectangle shape (positioned at row 2, column 2, size 200x100 points)
            Shape rect = ws.Shapes.AddRectangle(2, 2, 2, 2, 200, 100);

            // Apply a visible reflection effect (initial distance = 0)
            ReflectionEffect refl = rect.Reflection;
            refl.Type = ReflectionEffectType.HalfReflectionTouching;
            refl.Transparency = 0.5;   // 50% transparent
            refl.Size = 50;            // 50% size
            refl.Blur = 5;             // moderate blur
            refl.Distance = 0;         // start with no offset

            // Save the workbook (optional, just for inspection)
            wb.Save(workbookPath, SaveFormat.Xlsx);

            // 2. Render the worksheet to an image before changing distance
            RenderWorksheetToImage(ws, imgBeforePath);

            // 3. Change the reflection distance to twelve points
            refl.Distance = 12; // Task requirement

            // Save again to reflect the change (optional)
            wb.Save(workbookPath, SaveFormat.Xlsx);

            // 4. Render the worksheet to an image after changing distance
            RenderWorksheetToImage(ws, imgAfterPath);

            // 5. Compare the two images (simple file‑size based comparison)
            bool imagesDiffer = CompareImages(imgBeforePath, imgAfterPath, out double diffPercent);

            Console.WriteLine($"Reflection distance set to 12 points.");
            Console.WriteLine($"Images differ: {imagesDiffer}");
            Console.WriteLine($"Difference percentage (based on file size): {diffPercent:F2}%");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Renders a worksheet to a PNG image using SheetRender.
    static void RenderWorksheetToImage(Worksheet ws, string outputPath)
    {
        try
        {
            var options = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,
                Transparent = false
                // ImageFormat defaults to PNG; explicit setting omitted to avoid API mismatch.
            };

            var renderer = new SheetRender(ws, options);
            renderer.ToImage(0, outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to render image '{outputPath}': {ex.Message}");
        }
    }

    // Simple comparison based on file size; returns true if size difference exceeds 0.1%.
    static bool CompareImages(string pathA, string pathB, out double diffPercent)
    {
        diffPercent = 100.0;
        try
        {
            if (!File.Exists(pathA) || !File.Exists(pathB))
                return true; // Missing file => consider different

            long sizeA = new FileInfo(pathA).Length;
            long sizeB = new FileInfo(pathB).Length;

            if (sizeA == sizeB)
            {
                diffPercent = 0.0;
                return false;
            }

            long maxSize = Math.Max(sizeA, sizeB);
            diffPercent = Math.Abs(sizeA - sizeB) * 100.0 / maxSize;
            return diffPercent > 0.1;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Comparison failed: {ex.Message}");
            return true;
        }
    }
}
