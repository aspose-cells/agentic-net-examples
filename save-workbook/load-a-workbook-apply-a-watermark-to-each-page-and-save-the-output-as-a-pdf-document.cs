// Title: Apply a diagonal CONFIDENTIAL text watermark to each worksheet and save the workbook as a PDF with Aspose.Cells for .NET
// Common Searches: how to add a diagonal text watermark to all worksheets in an Excel file using Aspose.Cells C# | Aspose.Cells C# export workbook to PDF with custom watermark on each sheet | C# code to insert WordArt watermark into Excel worksheets before saving as PDF | apply CONFIDENTIAL watermark to an Excel workbook and convert to PDF using Aspose.Cells | set rotation angle for a shape watermark in Aspose.Cells C#
// Tags: add diagonal text watermark Aspose.Cells | worksheet TextEffect shape Aspose.Cells C# | export workbook to PDF with watermark C# | set shape rotation angle Aspose.Cells | apply WordArt watermark to Excel sheets

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an existing Excel workbook, iterates through each worksheet, and adds a WordArt TextEffect shape containing the text "CONFIDENTIAL" with a specified font, size, color, and a -45° rotation to create a diagonal watermark. After processing all sheets, it saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Watermark settings
            string watermarkText = "CONFIDENTIAL";
            string fontName = "Arial";
            int fontSize = 72;
            Color watermarkColor = Color.LightGray;
            double rotationAngle = -45;

            // Apply watermark to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Add WordArt shape (TextEffect) as watermark
                    Shape shape = sheet.Shapes.AddTextEffect(
                        MsoPresetTextEffect.TextEffect1,
                        watermarkText,
                        fontName,
                        fontSize,
                        false,
                        false,
                        0,   // left
                        0,   // top
                        0,   // width
                        0,   // height
                        0,   // textEffectFormat
                        0);  // textEffectShape

                    // Set visual properties
                    shape.Font.Color = watermarkColor;      // text color
                    shape.RotationAngle = rotationAngle;    // diagonal orientation
                    shape.Width = 500;                      // adjust as needed
                    shape.Height = 200;                     // adjust as needed
                    shape.ZOrderPosition = -1;              // send to back
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add watermark to sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Output PDF path
            string outputPath = "output.pdf";

            // Save workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
