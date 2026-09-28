// Title: Add a diagonal 'CONFIDENTIAL' watermark at coordinates (50,400) to every page of a PDF generated from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, inserts a semi‑transparent WordArt watermark reading 'CONFIDENTIAL' at position (50,400) on each worksheet, rotates it -45°, and saves the workbook as a PDF. | Show how to use Aspose.Cells' Shapes.AddTextEffect method to place a diagonal watermark at specific coordinates, set its transparency and Z‑order, then export the workbook to PDF in C#. | Provide a complete example that validates the input Excel file, adds the same cross‑string watermark to all sheets, and writes the output PDF with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add watermark text at exact X Y coordinates before PDF export | How to place a diagonal CONFIDENTIAL watermark on each page of a PDF generated from Excel using Aspose.Cells | C# Aspose.Cells add semi transparent WordArt shape as watermark to all worksheets | Set watermark rotation and transparency with Aspose.Cells when saving workbook as PDF | Apply identical watermark to every sheet in an Excel workbook and export to PDF with Aspose.Cells
// Tags: Aspose.Cells add watermark shape | C# Aspose.Cells PDF export with watermark | Shapes.AddTextEffect watermark coordinates | semi transparent rotated watermark Aspose.Cells | Excel to PDF watermark each worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an Excel file, adds a semi‑transparent diagonal 'CONFIDENTIAL' WordArt watermark at (50,400) on every worksheet, and saves the workbook as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(inputPath);

            // Add watermark text to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Insert a WordArt shape with the watermark text.
                // Parameters: preset, text, font name, font size, bold, italic,
                // left, top, width, height, leftColumn, topRow
                Shape watermark = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1, // preset effect
                    "CONFIDENTIAL",                  // watermark text
                    "Arial",                         // font name
                    48,                              // font size
                    false,                           // bold
                    false,                           // italic
                    50,                              // left (X) position
                    400,                             // top (Y) position
                    300,                             // width
                    100,                             // height
                    0,                               // left column
                    0);                              // top row

                // Make the watermark semi‑transparent
                watermark.Fill.Transparency = 0.5;

                // Remove outline by making the line fully transparent
                watermark.Line.Transparency = 1;

                // Rotate for typical watermark appearance
                watermark.RotationAngle = -45;

                // Send shape to back so it appears behind cell data
                watermark.ZOrderPosition = 0;
            }

            // Save the workbook as PDF (save rule)
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
