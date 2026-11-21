// Title: Convert an Excel workbook with multiple WordArt shapes to PDF and validate each gradient fill using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel WordArt shapes with gradient fill to PDF | How to inspect gradient fill of WordArt objects in a worksheet before PDF conversion using Aspose.Cells | Iterate over worksheet shapes in C# and log gradient type and angle with Aspose.Cells | Preserve individual WordArt gradient colors when converting Excel to PDF in .NET | Validate that each WordArt shape retains its gradient after saving workbook as PDF with Aspose.Cells
// Tags: Aspose.Cells WordArt gradient export | C# shape fill format inspection | preserve shape gradients during PDF conversion | validate individual WordArt gradients .NET | worksheet shapes iteration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // The example loads an Excel workbook, iterates through all shapes on the first worksheet, checks each shape's Fill.GradientColorType to determine if a gradient is applied, logs the gradient type and angle, and then saves the workbook as a PDF, ensuring the output directory exists.
    class WordArtGradientPdfConverter
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Get the fill format of the shape
                        FillFormat fill = shape.Fill;

                        // Determine if a gradient fill is applied
                        bool hasGradient = fill != null && fill.GradientColorType != GradientColorType.None;

                        Console.WriteLine($"Shape '{shape.Name}': Gradient applied = {hasGradient}");

                        // Print gradient details for manual inspection
                        if (hasGradient)
                        {
                            Console.WriteLine($"  Gradient Type: {fill.GradientColorType}");
                            // Use GradientDegree to represent the angle of the gradient
                            Console.WriteLine($"  Gradient Angle (Degree): {fill.GradientDegree}");
                        }
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Error processing shape '{shape.Name}': {exShape.Message}");
                    }
                }

                // Convert the workbook to PDF
                string outputPath = "output.pdf";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Conversion to PDF completed. File saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
