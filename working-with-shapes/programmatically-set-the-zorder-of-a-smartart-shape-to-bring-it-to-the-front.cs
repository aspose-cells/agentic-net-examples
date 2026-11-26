// Title: Programmatically bring a SmartArt shape to the front by setting its Z‑order in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to top layer | How to move a SmartArt object to front in an Excel file with Aspose.Cells | Increase Z-order of a worksheet shape using Aspose.Cells .NET | Bring Excel shape to front programmatically with Aspose.Cells API | Set maximum ZOrderPosition for a shape in a workbook using C#
// Tags: Aspose.Cells set ZOrderPosition | Excel SmartArt front layer Aspose.Cells | worksheet shape Z-order manipulation C# | increase shape Z-order Aspose.Cells | bring shape to top Excel .xlsx Aspose.Cells

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, checks for shapes on the first worksheet, calculates the highest current ZOrderPosition, assigns the first shape (e.g., a SmartArt object) a ZOrderPosition of max+1 to place it on the top layer, and saves the updated file.
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
            var workbook = new Workbook(inputPath);
            var worksheet = workbook.Worksheets[0];

            // Bring the first shape (if any) to the front
            if (worksheet.Shapes.Count > 0)
            {
                // Determine current maximum Z-order
                int maxZ = worksheet.Shapes.Max(s => s.ZOrderPosition);
                // Set the first shape's Z-order to be on top
                worksheet.Shapes[0].ZOrderPosition = maxZ + 1;
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
