// Title: Validate that every shape has a glow effect and add a default yellow glow before saving the workbook to XLSX with Aspose.Cells for .NET
// Common Searches: how to check and set a default glow effect for all shapes in an Excel workbook using Aspose.Cells C# | Aspose.Cells .NET iterate through worksheet shapes and apply missing glow before exporting to XLSX | C# code to ensure shape GlowEffect size is greater than zero in a workbook with Aspose.Cells | add semi‑transparent yellow glow to Excel shapes that lack glow using Aspose.Cells | validate shape formatting (glow) before saving workbook as XLSX in Aspose.Cells
// Tags: validate shape glow Aspose.Cells | apply default glow to Excel shapes C# | check GlowEffect size before saving workbook | Aspose.Cells shape formatting for XLSX export | add yellow semi‑transparent glow to shapes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeGlowValidator
{
    // The example loads an existing XLSX file, iterates through each worksheet and its shapes, checks each shape's GlowEffect size, and applies a default 5‑point semi‑transparent yellow glow when the size is zero. After processing all shapes, the workbook is saved as a new XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths – adjust as needed.
            string inputPath = "input.xlsx";
            string outputPath = "validated_output.xlsx";

            // Verify that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook \"{inputPath}\" loaded successfully.");

                // Iterate through all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the current worksheet.
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Access the GlowEffect object (always non‑null).
                        GlowEffect glow = shape.Glow;

                        // Determine if a glow is effectively defined.
                        // A glow is considered defined when its size is greater than zero.
                        bool hasGlow = glow.Size > 0;

                        if (!hasGlow)
                        {
                            // Apply a default glow: size 5 points, semi‑transparent yellow.
                            glow.Size = 5; // radius in points
                            CellsColor glowColor = workbook.CreateCellsColor();
                            glowColor.Color = System.Drawing.Color.Yellow;
                            glow.Color = glowColor;
                            glow.Transparency = 0.3; // 30% transparent

                            Console.WriteLine($"Applied default glow to shape \"{shape.Name}\" (type: {shape.Type}) on sheet \"{sheet.Name}\".");
                        }
                        else
                        {
                            Console.WriteLine($"Shape \"{shape.Name}\" already has a glow (size: {glow.Size}).");
                        }
                    }
                }

                // Save the validated workbook as XLSX.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Validated workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
