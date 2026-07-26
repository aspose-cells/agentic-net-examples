// Title: Aspose.Cells C# – Set Shape Reflection Distance to 12 pts and Detect Change via PNG Comparison
// Description: Creates two workbooks with identical rectangles, applies a half‑reflection effect, sets Reflection.Distance to 0 pts in the first file and to 12 pts in the second, exports each sheet to PNG, compares the images byte‑by‑byte to reveal visual differences, and reloads the 12‑point file to confirm the distance value is persisted.
// Keywords: Aspose.Cells | C# | shape reflection | Reflection.Distance | set reflection distance | PNG export | visual regression | byte comparison | Excel shape effect | .NET
// Common Searches: Aspose.Cells set shape reflection distance | C# compare exported Excel PNG images | how to verify reflection persistence Aspose.Cells | Reflection.Distance property example | automated visual test for Excel shape effects
// Developer Intent: Set a shape's reflection distance to 12 points and confirm the visual impact by comparing exported PNG snapshots.
// Use Cases: Generate baseline and modified workbooks to automate regression testing of shape reflection effects. | Validate that the Reflection.Distance setting is saved and correctly reloaded from an XLSX file. | Create documentation or UI previews showing different reflection distances by exporting PNG images. | Integrate byte‑level PNG comparison into CI pipelines for visual quality checks.
// AI Prompts: Write C# code with Aspose.Cells that adds a rectangle, applies a half‑reflection, sets Reflection.Distance to 12 points, and saves the workbook as XLSX and PNG. | Provide a C# method that loads two PNG files and returns the number of differing bytes for visual verification. | Show how to reload an Aspose.Cells workbook and read Shape.Reflection.Distance to ensure the value was persisted.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsReflectionDistanceDemo
{
    // Creates two workbooks with identical rectangles, applies a half‑reflection effect, sets Reflection.Distance to 0 pts in the first file and to 12 pts in the second, exports each sheet to PNG, compares the images byte‑by‑byte to reveal visual differences, and reloads the 12‑point file to confirm the distance value is persisted.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths for the generated files
                const string workbookPath0 = "ReflectionDistance0.xlsx";
                const string workbookPath12 = "ReflectionDistance12.xlsx";
                const string imagePath0 = "ReflectionDistance0.png";
                const string imagePath12 = "ReflectionDistance12.png";

                // -------------------------------------------------
                // 1. Create a workbook with default reflection distance (0)
                // -------------------------------------------------
                Workbook wbZero = new Workbook();
                Worksheet wsZero = wbZero.Worksheets[0];
                Shape shapeZero = wsZero.Shapes.AddRectangle(1, 1, 100, 100, 200, 150);
                shapeZero.Reflection.Type = ReflectionEffectType.HalfReflectionTouching;
                shapeZero.Reflection.Transparency = 0.5;
                shapeZero.Reflection.Size = 55;
                shapeZero.Reflection.Blur = 0.5;
                // Save workbook and export to PNG for visual comparison
                wbZero.Save(workbookPath0);
                wbZero.Save(imagePath0, SaveFormat.Png);

                // -------------------------------------------------
                // 2. Create a workbook with reflection distance set to 12 points
                // -------------------------------------------------
                Workbook wbTwelve = new Workbook();
                Worksheet wsTwelve = wbTwelve.Worksheets[0];
                Shape shapeTwelve = wsTwelve.Shapes.AddRectangle(1, 1, 100, 100, 200, 150);
                shapeTwelve.Reflection.Type = ReflectionEffectType.HalfReflectionTouching;
                shapeTwelve.Reflection.Transparency = 0.5;
                shapeTwelve.Reflection.Size = 55;
                shapeTwelve.Reflection.Blur = 0.5;
                shapeTwelve.Reflection.Distance = 12; // Set the required distance
                // Save workbook and export to PNG
                wbTwelve.Save(workbookPath12);
                wbTwelve.Save(imagePath12, SaveFormat.Png);

                // -------------------------------------------------
                // 3. Compare the two generated PNG files byte by byte
                // -------------------------------------------------
                if (!File.Exists(imagePath0) || !File.Exists(imagePath12))
                {
                    Console.WriteLine("One or both image files are missing; cannot compare.");
                }
                else
                {
                    byte[] bytesZero = File.ReadAllBytes(imagePath0);
                    byte[] bytesTwelve = File.ReadAllBytes(imagePath12);

                    int diffCount = 0;
                    int minLength = Math.Min(bytesZero.Length, bytesTwelve.Length);
                    for (int i = 0; i < minLength; i++)
                    {
                        if (bytesZero[i] != bytesTwelve[i])
                            diffCount++;
                    }
                    diffCount += Math.Abs(bytesZero.Length - bytesTwelve.Length);

                    Console.WriteLine($"Total differing bytes: {diffCount}");
                    Console.WriteLine(diffCount > 0
                        ? "Visual change detected due to reflection distance = 12 points."
                        : "No visual change detected.");
                }

                // -------------------------------------------------
                // 4. Verify persistence by reloading the workbook with distance = 12
                // -------------------------------------------------
                if (File.Exists(workbookPath12))
                {
                    Workbook loaded = new Workbook(workbookPath12);
                    Shape loadedShape = loaded.Worksheets[0].Shapes[0];
                    double loadedDistance = loadedShape.Reflection.Distance;
                    Console.WriteLine($"Loaded reflection distance: {loadedDistance} points");
                }
                else
                {
                    Console.WriteLine($"Workbook file '{workbookPath12}' not found; cannot verify persistence.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
