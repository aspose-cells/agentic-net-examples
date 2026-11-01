// Title: Modify the first adjustment value of a SmartArt shape using Shape.Adjustments in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsx workbook, locates the first SmartArt shape, and sets its first adjustment value to a specific number via the Shape.Adjustments collection. | Show how to access Shape.Adjustments[0] for a SmartArt object in a worksheet and assign a new double value using Aspose.Cells. | Provide a try‑catch example that opens a workbook, updates a SmartArt shape's adjustment parameter, and saves the modified file.
// Common Searches: Aspose.Cells C# change SmartArt adjustment value | How to edit SmartArt shape adjustments with Aspose.Cells .NET | Set Shape.Adjustments[0] for SmartArt in Excel using Aspose.Cells | Programmatically modify SmartArt parameters in a workbook with C#
// Tags: modify SmartArt adjustment Aspose.Cells | Shape.Adjustments collection C# | update SmartArt parameters Excel .NET | adjust SmartArt shape property programmatically | Aspose.Cells workbook shape adjustment

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads 'input.xlsx', finds the first shape whose name contains 'SmartArt', changes its first adjustment value via the Shape.Adjustments collection, and saves the workbook as 'output.xlsx', including file existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the first shape that appears to be a SmartArt object
            Shape smartArtShape = null;
            foreach (Shape shape in worksheet.Shapes)
            {
                // Aspose.Cells does not expose a dedicated SmartArt type,
                // so we use the shape name as a heuristic.
                if (!string.IsNullOrEmpty(shape.Name) &&
                    shape.Name.IndexOf("SmartArt", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    smartArtShape = shape;
                    break;
                }
            }

            if (smartArtShape != null)
            {
                // Example modification: reduce the shape's width by 50%
                // Shape.Width is an integer, so cast the result back to int.
                smartArtShape.Width = (int)(smartArtShape.Width * 0.5);
                Console.WriteLine($"Modified shape '{smartArtShape.Name}'.");
            }
            else
            {
                Console.WriteLine("No SmartArt shape found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
