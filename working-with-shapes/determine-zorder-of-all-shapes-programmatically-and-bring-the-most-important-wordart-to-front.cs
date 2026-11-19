// Title: Programmatically find the top‑most WordArt shape by Z‑order and move it to the front using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells get ZOrderPosition of shapes in a worksheet | How to bring a specific WordArt shape to the front in an Excel file using Aspose.Cells | Set maximum Z‑order for a shape with Aspose.Cells .NET API | Identify shape with highest Z‑order in an Excel workbook using Aspose.Cells | Programmatically move WordArt to front of other shapes in Excel via Aspose.Cells
// Tags: retrieve shape ZOrderPosition Aspose.Cells .NET | set shape ZOrderPosition to maximum Aspose.Cells | move WordArt to front Excel Aspose.Cells | iterate worksheet shapes Aspose.Cells | find topmost WordArt shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code loads an Excel workbook, iterates all shapes on the first worksheet to locate the WordArt shape with the highest ZOrderPosition, updates that shape's ZOrderPosition to the maximum count to bring it to the front, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Find the WordArt shape with the highest Z‑order
            Shape mostImportantWordArt = null;
            int highestZ = -1;

            foreach (Shape shape in sheet.Shapes)
            {
                // WordArt shapes are identified by the IsWordArt property
                if (shape.IsWordArt)
                {
                    int zOrder = shape.ZOrderPosition;
                    if (zOrder > highestZ)
                    {
                        highestZ = zOrder;
                        mostImportantWordArt = shape;
                    }
                }
            }

            // Bring the identified WordArt to the front by setting its Z‑order to the maximum
            if (mostImportantWordArt != null)
            {
                try
                {
                    int maxZ = sheet.Shapes.Count; // highest possible Z‑order after adding a shape
                    mostImportantWordArt.ZOrderPosition = maxZ;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to adjust Z‑order: {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
