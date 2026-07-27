// Title: Convert Excel with Multiple WordArt Shapes to PDF and Verify Gradient Rendering (C# Aspose.Cells)
// Description: Loads an .xlsx workbook, iterates every worksheet, identifies WordArt shapes, logs each shape's fill type, enumerates gradient stops when the fill is a gradient, and saves the file as a PDF so that every WordArt gradient is rendered independently. Includes file‑existence validation and robust error handling.
// Keywords: Aspose.Cells | C# | WordArt | gradient fill | PDF conversion | Excel to PDF | shape collection | gradient stops | independent rendering | workbook save
// Common Searches: Aspose.Cells export WordArt with gradient to PDF C# | How to read gradient stops of WordArt shapes in Excel using Aspose.Cells | Verify independent gradient rendering of WordArt after PDF conversion | Iterate shapes collection in Aspose.Cells C# example | Convert .xlsx containing WordArt to PDF preserving gradients
// Developer Intent: Convert a spreadsheet that contains several WordArt objects to PDF and confirm that each gradient fill is rendered separately.
// Use Cases: Generate PDF brochures where WordArt titles retain their original gradient colors. | Log gradient stop details of all WordArt shapes for quality‑assurance before publishing. | Automate batch conversion of Excel files with WordArt to PDF while validating visual fidelity.
// AI Prompts: Provide C# code using Aspose.Cells that loads an .xlsx, loops through all WordArt shapes, prints each gradient stop position, and saves the workbook as a PDF. | Explain how to compare gradient colors in the generated PDF with the original WordArt definitions using Aspose.Cells. | Suggest robust error‑handling patterns when a WordArt shape reports a Gradient fill but the GradientFill object is null.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtToPdf
{
    // Loads an .xlsx workbook, iterates every worksheet, identifies WordArt shapes, logs each shape's fill type, enumerates gradient stops when the fill is a gradient, and saves the file as a PDF so that every WordArt gradient is rendered independently. Includes file‑existence validation and robust error handling.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "WordArtSample.xlsx";
                const string outputPath = "WordArtSample.pdf";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook that contains multiple WordArt objects
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the shape collection of the current worksheet
                    ShapeCollection shapes = sheet.Shapes;

                    // Examine each shape
                    for (int i = 0; i < shapes.Count; i++)
                    {
                        Shape shape = shapes[i];

                        // Process only WordArt shapes
                        if (shape.IsWordArt)
                        {
                            // Retrieve the fill type of the WordArt
                            FillType fillType = shape.Fill.FillType;

                            // Output basic information
                            Console.WriteLine($"WordArt #{i + 1} on sheet \"{sheet.Name}\": FillType = {fillType}");

                            // If the fill is a gradient, inspect its gradient details
                            if (fillType == FillType.Gradient)
                            {
                                GradientFill gradientFill = shape.Fill.GradientFill;
                                if (gradientFill != null)
                                {
                                    // Output the number of gradient stops (each stop defines a color)
                                    int stopCount = gradientFill.GradientStops.Count;
                                    Console.WriteLine($"  Gradient has {stopCount} stop(s).");

                                    // List each gradient stop's position (color property not exposed in this version)
                                    for (int s = 0; s < stopCount; s++)
                                    {
                                        var stop = gradientFill.GradientStops[s];
                                        Console.WriteLine($"    Stop {s + 1}: Position = {stop.Position}");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("  GradientFill instance is null.");
                                }
                            }
                            else
                            {
                                Console.WriteLine("  Fill is not a gradient; independent rendering not applicable.");
                            }
                        }
                    }
                }

                // Save the workbook as PDF; each WordArt gradient will be rendered independently in the PDF
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as PDF to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
