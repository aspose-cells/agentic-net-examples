// Title: Validate Aspose.Cells GradientFill stops against a CSS two‑color gradient (C#)
// Description: This C# example creates a workbook, adds a rectangle shape, applies a horizontal two‑color gradient (red → blue) with Aspose.Cells, extracts the GradientStopCollection, and checks each stop’s position, ARGB color and opacity against the expected CSS linear‑gradient values. Mismatches are reported and the workbook is saved.
// Keywords: Aspose.Cells | GradientFill | GradientStopCollection | C# gradient validation | CSS linear-gradient | color stop verification | Excel shape fill | gradient opacity | automated testing | WordArt gradient
// Common Searches: Aspose.Cells verify gradient stops | C# compare gradient fill with CSS | Check gradient stop positions Aspose.Cells | Validate gradient opacity in Excel shape | How to extract CSS linear-gradient from Aspose.Cells
// Developer Intent: Confirm that the gradient fill produced by Aspose.Cells exactly matches the specified CSS two‑color gradient.
// Use Cases: Automated quality‑gate to ensure Excel shapes follow design‑system gradient specs before distribution. | Unit‑test suite for WordArt‑to‑Excel conversions that validates color fidelity and transparency. | Diagnostic tool for troubleshooting mismatched gradient stops when importing external CSS gradients into a workbook.
// AI Prompts: Generate a C# method that receives a GradientStopCollection and an array of expected (position, Color) tuples, compares them with a tolerance, and returns a detailed validation report. | Create code that extracts gradient stop positions, colors, and transparency from an Aspose.Cells shape and builds the equivalent CSS linear‑gradient string. | Write an NUnit test that builds a two‑color gradient using SetTwoColorGradient and asserts that the resulting GradientStopCollection contains the correct stops, colors and opacity.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGradientValidation
{
    // This C# example creates a workbook, adds a rectangle shape, applies a horizontal two‑color gradient (red → blue) with Aspose.Cells, extracts the GradientStopCollection, and checks each stop’s position, ARGB color and opacity against the expected CSS linear‑gradient values. Mismatches are reported and the workbook is saved.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape (used as a placeholder for WordArt)
                Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);

                // Set the fill type to gradient
                shape.Fill.FillType = FillType.Gradient;

                // Apply a two‑color gradient (first color: Red, second color: Blue)
                // Gradient style: Horizontal, variant: 1
                shape.Fill.GradientFill.SetTwoColorGradient(
                    Color.Red,
                    Color.Blue,
                    GradientStyleType.Horizontal,
                    1);

                // Retrieve the gradient stops created by the above method
                GradientStopCollection stops = shape.Fill.GradientFill.GradientStops;

                // Expected gradient stops for a two‑color gradient:
                // Position 0.0 -> Red (opaque)
                // Position 1.0 -> Blue (opaque)
                var expectedStops = new (double Position, Color Color)[]
                {
                    (0.0, Color.Red),
                    (1.0, Color.Blue)
                };

                bool validationPassed = true;

                // Validate count
                if (stops.Count != expectedStops.Length)
                {
                    validationPassed = false;
                    Console.WriteLine($"Validation failed: Expected {expectedStops.Length} stops, found {stops.Count}.");
                }
                else
                {
                    // Compare each stop
                    for (int i = 0; i < stops.Count; i++)
                    {
                        GradientStop actual = stops[i];
                        double expectedPos = expectedStops[i].Position;
                        Color expectedColor = expectedStops[i].Color;

                        // Position comparison (allow small tolerance)
                        if (Math.Abs(actual.Position - expectedPos) > 0.0001)
                        {
                            validationPassed = false;
                            Console.WriteLine($"Stop {i}: Position mismatch. Expected {expectedPos}, got {actual.Position}");
                        }

                        // Color comparison
                        Color actualColor = actual.CellsColor.Color;
                        if (actualColor.ToArgb() != expectedColor.ToArgb())
                        {
                            validationPassed = false;
                            Console.WriteLine($"Stop {i}: Color mismatch. Expected {expectedColor}, got {actualColor}");
                        }

                        // Alpha (transparency) comparison – CellsColor.Transparency is 0‑255 where 0 = opaque
                        // For this example we used opaque colors, so expected alpha = 255
                        int expectedAlpha = 255;
                        int actualAlpha = 255 - (int)actual.CellsColor.Transparency; // Cast to int if Transparency is double
                        if (actualAlpha != expectedAlpha)
                        {
                            validationPassed = false;
                            Console.WriteLine($"Stop {i}: Alpha mismatch. Expected {expectedAlpha}, got {actualAlpha}");
                        }
                    }
                }

                Console.WriteLine(validationPassed
                    ? "All gradient stops match the expected CSS definition."
                    : "Gradient stop validation failed.");

                // Save the workbook (lifecycle rule: save)
                workbook.Save("GradientValidationResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
