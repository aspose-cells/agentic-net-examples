// Title: Add a semi‑transparent diagonal WordArt watermark to each worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add diagonal WordArt watermark to all worksheets | how to set semi transparent fill and line for a WordArt shape in Excel using Aspose.Cells | iterate through worksheets and insert free‑floating shape behind cells with Aspose.Cells | create confidential CONFIDENTIAL watermark in Excel workbook using Aspose.Cells C#
// Tags: add wordart shape watermark aspnet cells c# | set shape fill and line transparency aspnet cells | rotate shape -45 degrees aspnet cells c# | iterate worksheets add shape aspnet cells | freefloating shape placement behind cells aspnet cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an existing Excel workbook, loops through every worksheet, inserts a WordArt text effect shape containing the text "CONFIDENTIAL", rotates it -45 degrees to make it diagonal, applies 50% transparency to both fill and line, sets the shape as free‑floating and positions it behind other objects, then saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add a WordArt shape that will serve as the watermark
                // Provide all required parameters, including upper-left row and column
                Shape watermark = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1, // preset text effect
                    "CONFIDENTIAL",                  // watermark text
                    "Arial",                         // font name
                    36,                              // font size
                    false,                           // bold
                    false,                           // italic
                    0,                               // left position
                    0,                               // top position
                    400,                             // width
                    100,                             // height
                    0,                               // upper left row
                    0);                              // upper left column

                // Rotate to make it diagonal
                watermark.RotationAngle = -45;

                // Make the shape free‑floating so it stays on the sheet
                watermark.Placement = PlacementType.FreeFloating;

                // Set semi‑transparent fill and line
                watermark.Fill.Transparency = 0.5; // 50% transparent fill
                watermark.Line.Transparency = 0.5; // 50% transparent line

                // Place the shape behind other objects by setting a low Z‑order position
                watermark.ZOrderPosition = 0;
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
