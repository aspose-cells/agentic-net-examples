// Title: Load a password‑protected Excel file, add a diagonal semi‑transparent CONFIDENTIAL watermark, and save as PDF with Aspose.Cells for .NET
// Common Searches: how to open an encrypted .xlsx with a password and export to PDF using Aspose.Cells in C# | add a diagonal CONFIDENTIAL watermark to each worksheet before converting Excel to PDF with Aspose.Cells | Aspose.Cells C# example for loading a password‑protected workbook and applying a WordArt watermark | convert password protected Excel to PDF with semi‑transparent watermark using Aspose.Cells .NET | C# code to overlay a rotated text shape on all sheets of an encrypted workbook and save as PDF
// Tags: load encrypted workbook with password Aspose.Cells | add diagonal text watermark using WordArt Aspose.Cells | convert Excel to PDF with overlay watermark Aspose.Cells | apply semi‑transparent shape to worksheets C# | save workbook as PDF Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example checks for an encrypted Excel file, loads it with the provided password via LoadOptions, adds a semi‑transparent red WordArt shape titled "CONFIDENTIAL" rotated -45° on every worksheet, and then saves the workbook as a PDF document.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "EncryptedInput.xlsx";
            const string outputPath = "OutputWithWatermark.pdf";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the encrypted workbook with the password
            var loadOptions = new LoadOptions
            {
                Password = "yourPassword" // replace with actual password
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Add a text watermark to each worksheet using a WordArt shape
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Create a WordArt shape with the watermark text
                var shape = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1, // preset effect
                    "CONFIDENTIAL",                  // text
                    "Arial",                         // font name
                    72,                              // font size
                    false,                           // bold
                    false,                           // italic
                    0,                               // left (pixels)
                    0,                               // top (pixels)
                    500,                             // width (pixels)
                    200,                             // height (pixels)
                    0,                               // text effect (default)
                    0);                              // text effect color (default)

                // Configure appearance
                shape.Font.Color = Color.FromArgb(128, Color.Red); // semi‑transparent red
                shape.Font.IsBold = true;
                shape.RotationAngle = -45;                         // diagonal orientation
                shape.Placement = PlacementType.FreeFloating;      // overlay, not background
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
