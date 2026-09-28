// Title: How to add a non‑rotated 20% opacity text watermark to every worksheet in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that inserts a plain text watermark "Confidential" on each worksheet, sets RotationAngle = 0 and Fill.Transparency = 0.8, then saves the workbook. | Show a C# example that iterates through all worksheets in a .xlsx file and adds a text effect shape as a watermark with zero rotation and 80% transparency using Aspose.Cells. | Provide a C# snippet that creates a non‑rotated text watermark with 20% opacity on every sheet of an Excel workbook and saves the result with Aspose.Cells.
// Common Searches: Aspose.Cells add text watermark to all worksheets C# | C# set watermark rotation angle to zero in Excel using Aspose.Cells | How to apply 20% opacity text watermark in an Excel workbook with Aspose.Cells | Programmatically add plain text effect shape as watermark in .xlsx using Aspose.Cells | Aspose.Cells watermark each sheet without rotation
// Tags: Aspose.Cells text effect shape watermark | C# add non‑rotated watermark to Excel worksheets | Excel workbook 20% opacity watermark Aspose.Cells | iterate worksheets add watermark programmatically | set shape transparency Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing XLSX file, loops through every worksheet, and adds a plain text effect shape containing "Confidential". The shape is configured with zero rotation and 80% fill and line transparency (equivalent to 20% opacity), placed to move with cells, unlocked, and sent to the back before the workbook is saved to a new file.
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

            // Apply a text watermark to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Add a plain text effect shape.
                    // The overload requires upper‑left and lower‑right cell positions plus height and width.
                    Shape watermarkShape = sheet.Shapes.AddTextEffect(
                        (MsoPresetTextEffect)0,          // Plain text effect
                        "Confidential",                  // Watermark text
                        "Arial",                         // Font name
                        48,                              // Font size
                        false,                           // Bold
                        false,                           // Italic
                        0, 0,                            // Upper‑left cell (row, column)
                        0, 0,                            // Lower‑right cell (row, column)
                        0, 0);                           // Height, Width (auto‑size)

                    // Configure shape appearance
                    watermarkShape.RotationAngle = 0;                     // No rotation
                    watermarkShape.Fill.Transparency = 0.8;              // 80% transparent fill
                    watermarkShape.Line.Transparency = 0.8;              // 80% transparent line
                    watermarkShape.Placement = PlacementType.Move;      // Move with cells
                    watermarkShape.IsLocked = false;                     // Allow editing
                    watermarkShape.ZOrderPosition = 0;                   // Send to back
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add watermark to sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the watermark applied
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with watermark to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
