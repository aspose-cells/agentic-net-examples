// Title: Add margin offsets to the first shape's absolute position in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Read the Top and Left values of the first worksheet shape, add a 10‑point vertical and 15‑point horizontal margin, and save the workbook with Aspose.Cells in C#. | Programmatically shift a drawing object in an Excel file by a custom point offset using the Aspose.Cells Shape.Top and Shape.Left properties. | Update the absolute coordinates of an Excel shape to include a margin, then persist the changes to a new .xlsx file via Aspose.Cells.
// Common Searches: how to move an Excel shape by points using Aspose.Cells C# | retrieve shape coordinates and apply offset in .NET workbook | add custom margin to first drawing object in Aspose.Cells worksheet | adjust shape top left position programmatically in Excel file | Aspose.Cells shape reposition example with margin offsets
// Tags: Aspose.Cells shape position offset | C# adjust Excel drawing coordinates | apply point margin to worksheet shape | modify absolute position of Excel shape | move first shape with Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook, accesses the first worksheet, retrieves the first shape's Top and Left coordinates, adds a 10‑point vertical and 15‑point horizontal margin, updates the shape's absolute position, and saves the modified file to a new location.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure at least one shape exists
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape
            Shape shape = worksheet.Shapes[0];

            // Current absolute position (points)
            int currentTop = shape.Top;
            int currentLeft = shape.Left;

            // Margin offsets (points)
            int marginTop = 10;   // move down by 10 points
            int marginLeft = 15;  // move right by 15 points

            // Apply the margin offsets
            shape.Top = currentTop + marginTop;
            shape.Left = currentLeft + marginLeft;

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
