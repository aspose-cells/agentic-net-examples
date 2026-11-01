// Title: Apply a diagonal CONFIDENTIAL text‑effect watermark to all worksheets and export the workbook to PDF with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add rotated text watermark to each worksheet before PDF export | How to create a semi‑transparent WordArt watermark in an Excel file using Aspose.Cells | Export Excel workbook to PDF with diagonal watermark using Aspose.Cells .NET API | Add watermark shape to all sheets in a workbook with Aspose.Cells and save as PDF
// Tags: Aspose.Cells text effect watermark | rotate watermark shape worksheet | semi transparent shape fill Aspose.Cells | save workbook as PDF with watermark | apply watermark to all worksheets C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program checks for the input XLSX file, loads it into an Aspose.Cells Workbook, iterates through every worksheet, and adds a WordArt‑style text effect shape containing the word "CONFIDENTIAL". Each shape is positioned centrally, set to free‑floating placement, rotated -45°, and given 50 % transparency for both fill and outline. After processing all sheets, the workbook is saved as a PDF document.
class WatermarkPdfExport
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            const string watermarkText = "CONFIDENTIAL";

            // Add watermark to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Create a WordArt‑like text effect shape
                    Shape watermarkShape = sheet.Shapes.AddTextEffect(
                        MsoPresetTextEffect.TextEffect1,
                        watermarkText,
                        "Arial",
                        72,          // font size
                        false,       // bold
                        false,       // italic
                        0, 0,        // left, top (initial)
                        500, 200,    // width, height
                        0, 0);       // additional required parameters (shape type placeholders)

                    // Configure appearance
                    watermarkShape.Placement = PlacementType.FreeFloating;
                    watermarkShape.RotationAngle = -45; // rotate the text
                    // Semi‑transparent fill and outline
                    watermarkShape.Fill.Transparency = 0.5; // 50 % transparent fill
                    watermarkShape.Line.Transparency = 0.5; // 50 % transparent outline

                    // Position the shape roughly at the centre of the sheet
                    watermarkShape.Top = 200;
                    watermarkShape.Left = 100;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add watermark to sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // Save as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
