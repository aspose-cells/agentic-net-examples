// Title: Add a Rotated Semi‑Transparent WordArt Watermark to Every Worksheet (Aspose.Cells C#)
// Description: C# example that creates or loads an Aspose.Cells workbook, loops through all worksheets, inserts a WordArt shape with custom text, sets 50 % transparency, rotates it -45°, customizes the font, and saves the file. Demonstrates using the ShapeCollection, TextEffectFormat, and file‑system checks.
// Keywords: Aspose.Cells watermark | C# WordArt shape | rotate WordArt Excel | transparent watermark Aspose.Cells | iterate worksheets add shape | Excel .NET shape collection | programmatic watermark Excel | Aspose.Cells C# example
// Common Searches: how to add a rotated watermark to all sheets with Aspose.Cells .NET | Aspose.Cells C# set WordArt transparency | apply diagonal CONFIDENTIAL watermark to every worksheet | add WordArt shape to Excel workbook using Aspose.Cells | save workbook after inserting watermarks in C#
// Developer Intent: Programmatically place a semi‑transparent, diagonal WordArt watermark on each worksheet of an Excel file.
// Use Cases: Mark confidential reports with a diagonal "CONFIDENTIAL" label across all tabs. | Brand internal templates by overlaying a faint company name or logo on every sheet. | Insert a draft watermark before distributing generated workbooks to reviewers.
// AI Prompts: Generate C# code with Aspose.Cells that adds a semi‑transparent rotated WordArt watermark to all worksheets. | Show how to change the transparency, rotation angle, and font of a WordArt shape in an Excel workbook using Aspose.Cells. | Provide an example that creates the output folder if it does not exist before saving the watermarked workbook.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates or loads an Aspose.Cells workbook, loops through all worksheets, inserts a WordArt shape with custom text, sets 50 % transparency, rotates it -45°, customizes the font, and saves the file. Demonstrates using the ShapeCollection, TextEffectFormat, and file‑system checks.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the shape collection of the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Add a WordArt shape that will serve as the watermark
                // Parameters: style, text, topRow, top offset, leftColumn, left offset, height, width
                Shape wordArt = shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,   // preset style
                    "CONFIDENTIAL",                     // watermark text
                    0, 0,                               // top row and vertical offset (pixels)
                    0, 0,                               // left column and horizontal offset (pixels)
                    500, 800);                          // height and width (pixels)

                // Make the WordArt semi‑transparent
                wordArt.Fill.Transparency = 0.5; // 50% transparency

                // Rotate the WordArt to give a typical watermark appearance
                wordArt.RotationAngle = -45; // degrees

                // Optional: customize the text appearance
                if (wordArt.IsWordArt)
                {
                    TextEffectFormat textEffect = wordArt.TextEffect;
                    textEffect.FontBold = true;
                    textEffect.FontItalic = false;
                    textEffect.FontName = "Arial";
                    textEffect.FontSize = 48;
                    // FontColor property is not available in some versions; omitted for compatibility
                }
            }

            // Define output file path
            string outputPath = "WatermarkedWorkbook.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the watermarks applied
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
