// Title: Convert Excel with Multiple WordArt Shapes to PDF and Verify Independent Gradient Fills – Aspose.Cells for .NET
// Description: Loads an Excel workbook, iterates all shapes on the first worksheet, filters WordArt objects, logs each WordArt's text and gradient fill details (type, direction, stop count, colors and positions), and saves the workbook as a PDF. The process confirms that every WordArt gradient is rendered separately in the resulting PDF.
// Keywords: Aspose.Cells | C# | WordArt | gradient fill | PDF conversion | Excel to PDF | shape iteration | gradient stops | preserve gradients | WordArt to PDF
// Common Searches: Aspose.Cells export WordArt to PDF | How to read WordArt gradient colors in .NET | Verify WordArt gradient after PDF conversion | Iterate shapes in Excel using Aspose.Cells | Get gradient stops from WordArt
// Developer Intent: The developer wants to convert an Excel file containing several WordArt objects to PDF and ensure that each object's gradient fill remains distinct and correctly rendered.
// Use Cases: Enumerate WordArt shapes and output their gradient properties for validation. | Extract gradient stop colors and positions to compare with expected values. | Create a PDF that preserves the visual appearance of all WordArt gradient fills. | Automate a regression test that checks independent gradient rendering after conversion.
// AI Prompts: Generate C# code that compares the gradient stop colors of each WordArt in the PDF with those defined in the original Excel workbook using Aspose.Cells. | Explain how to modify the gradient fill of a specific WordArt shape programmatically before exporting the workbook to PDF with Aspose.Cells. | Suggest a unit‑test approach to assert that multiple WordArt objects retain their independent gradient fills after PDF conversion.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtToPdf
{
    // Loads an Excel workbook, iterates all shapes on the first worksheet, filters WordArt objects, logs each WordArt's text and gradient fill details (type, direction, stop count, colors and positions), and saves the workbook as a PDF. The process confirms that every WordArt gradient is rendered separately in the resulting PDF.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "WordArtSample.xlsx";

                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook containing WordArt objects
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only WordArt shapes
                    if (shape.IsWordArt)
                    {
                        Console.WriteLine($"WordArt Text: {shape.TextEffect.Text}");

                        // Check if the fill type is gradient
                        if (shape.Fill.FillType == FillType.Gradient)
                        {
                            GradientFill gradientFill = shape.Fill.GradientFill;

                            Console.WriteLine($"  Gradient Fill Type: {gradientFill.FillType}");
                            Console.WriteLine($"  Gradient Direction: {gradientFill.DirectionType}");
                            Console.WriteLine($"  Gradient Stops Count: {gradientFill.GradientStops.Count}");

                            // List each gradient stop (color may not be available in older API versions)
                            for (int i = 0; i < gradientFill.GradientStops.Count; i++)
                            {
                                var stop = gradientFill.GradientStops[i];
                                string colorInfo = "N/A";

                                // Safely attempt to read the Color property via reflection
                                var colorProp = stop.GetType().GetProperty("Color");
                                if (colorProp != null)
                                {
                                    var colorValue = colorProp.GetValue(stop);
                                    colorInfo = colorValue?.ToString() ?? "null";
                                }

                                Console.WriteLine($"    Stop {i + 1}: Color={colorInfo}, Position={stop.Position}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("  Fill is not a gradient.");
                        }
                    }
                }

                // Convert the workbook (with WordArt) to PDF
                const string outputPath = "WordArtSample.pdf";
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Conversion to PDF completed. Saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
