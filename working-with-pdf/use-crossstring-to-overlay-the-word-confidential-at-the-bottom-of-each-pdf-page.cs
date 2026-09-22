// Title: Add a red semi‑transparent “Confidential” text watermark to the bottom of each worksheet in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that inserts a red, 50 % transparent WordArt shape containing the word “Confidential” at the bottom‑center of every worksheet and saves the workbook. | Show how to programmatically apply a horizontal text‑effect watermark to all sheets of an XLSX file, configuring font size, color, transparency, and placement with Aspose.Cells in .NET.
// Common Searches: Aspose.Cells C# add confidential watermark to every sheet of an Excel file | how to place semi transparent text at the bottom of Excel worksheets using Aspose.Cells | C# programmatically overlay WordArt watermark on all worksheets in XLSX | set watermark transparency and position with Aspose.Cells .NET | add bottom‑center text watermark to Excel workbook via code
// Tags: Aspose.Cells add WordArt watermark to worksheets | C# Excel worksheet bottom text overlay | semi transparent text effect Aspose.Cells | programmatic Excel watermark placement | confidential watermark Excel file .NET | Aspose.Cells shape transparency settings

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing XLSX file, iterates through each worksheet, and adds a WordArt shape that displays the word “Confidential” in red, bold 36‑point font with 50 % fill and line transparency. The shape is positioned at the bottom‑center of the page based on the worksheet's page dimensions, then the modified workbook is saved to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string watermarkText = "Confidential";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply watermark to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Add a WordArt shape as a watermark
                    // Parameters: preset effect, text, font name, size, bold, italic,
                    // left, top, width, height, textEffect, shapeId
                    Shape watermarkShape = sheet.Shapes.AddTextEffect(
                        MsoPresetTextEffect.TextEffect1,
                        watermarkText,
                        "Arial",
                        36,
                        true,
                        false,
                        0,
                        0,
                        500,
                        100,
                        0,
                        0);

                    // Set appearance
                    watermarkShape.Fill.Transparency = 0.5;          // Semi‑transparent fill
                    watermarkShape.Line.Transparency = 0.5;          // Semi‑transparent outline
                    watermarkShape.Font.Color = Color.Red;
                    watermarkShape.Font.IsBold = true;
                    watermarkShape.Font.Size = 36;
                    watermarkShape.RotationAngle = 0;                // Horizontal orientation
                    watermarkShape.Placement = PlacementType.FreeFloating;
                    watermarkShape.Name = "Watermark";

                    // Center the shape on the page (approximate)
                    double pageWidth = sheet.PageSetup.PaperWidth;
                    double pageHeight = sheet.PageSetup.PaperHeight;
                    watermarkShape.Left = (int)((pageWidth - watermarkShape.Width) / 2);
                    watermarkShape.Top = (int)((pageHeight - watermarkShape.Height) / 2);
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Failed to add watermark to sheet '{sheet.Name}': {exShape.Message}");
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Watermarked workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
