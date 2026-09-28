// Title: Create a diagonal semi‑transparent 'Confidential' text watermark on every worksheet of an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, adds a 45° rotated, 50% transparent text shape reading 'Confidential' to each worksheet, and saves the workbook. | Show how to configure a TextEffect shape in Aspose.Cells to set font, size, rotation, transparency, and hide its border for a watermark applied across all sheets. | Provide a reusable C# method that iterates through a Workbook’s worksheets and inserts a free‑floating text watermark shape with custom appearance using Aspose.Cells.
// Common Searches: Aspose.Cells C# add diagonal confidential watermark to all Excel worksheets | How to programmatically place a semi‑transparent text overlay on each sheet in an .xlsx using Aspose.Cells | C# example for rotating and setting transparency of a shape in Aspose.Cells | Apply the same text watermark to multiple worksheets with Aspose.Cells for .NET | Hide shape border while adding watermark in Aspose.Cells workbook
// Tags: Aspose.Cells add diagonal text watermark | C# set shape transparency Aspose.Cells | Aspose.Cells rotate text shape worksheet | free‑floating shape placement Aspose.Cells | hide shape border Aspose.Cells | apply watermark to all worksheets C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing .xlsx workbook, iterates through each worksheet, and inserts a free‑floating TextEffect shape containing the word 'Confidential'. The shape is styled with Arial 72 pt, 45° rotation, 50% transparency, and its border line weight is set to zero, creating a diagonal semi‑transparent watermark on every sheet. After processing, the workbook is saved to the specified output path.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply a text watermark to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Add a text effect shape that will serve as the watermark
                    Shape watermarkShape = sheet.Shapes.AddTextEffect(
                        MsoPresetTextEffect.TextEffect1,
                        "Confidential",
                        "Arial",
                        72,
                        true,
                        false,
                        0,
                        0,
                        500,
                        200,
                        0,
                        0);

                    // Set shape properties to mimic a watermark
                    watermarkShape.Fill.Transparency = 0.5;          // 0 = opaque, 1 = fully transparent
                    watermarkShape.RotationAngle = 45;              // diagonal orientation
                    watermarkShape.Placement = PlacementType.FreeFloating;
                    // Hide the shape border (line). In Aspose.Cells LineFormat does not expose IsVisible, so set weight to 0.
                    watermarkShape.Line.Weight = 0;
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Failed to add watermark to sheet '{sheet.Name}': {exShape.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Watermarked workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
