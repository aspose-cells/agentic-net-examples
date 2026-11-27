// Title: Create and save an XLSX workbook with a WordArt CONFIDENTIAL watermark while retaining shape properties using Aspose.Cells for .NET
// Common Searches: how to add a WordArt watermark to an Excel worksheet with Aspose.Cells C# | preserve WordArt shape transparency and rotation when saving workbook as XLSX using Aspose.Cells | lock WordArt shape movement and selection in Aspose.Cells before exporting to XLSX | set WordArt fill color and hide outline line in Aspose.Cells C# example | save workbook with WordArt watermark to XLSX preserving all shape attributes Aspose.Cells
// Tags: AddWordArt watermark Aspose.Cells | WordArt shape fill transparency Aspose.Cells | WordArt rotation Aspose.Cells | WordArt shape lock properties Aspose.Cells | save workbook as XLSX preserving shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtWatermarkSave
{
    // The example creates a new Workbook, inserts a WordArt shape with the text "CONFIDENTIAL" as a diagonal watermark, configures its fill color, transparency, hides the outline, rotates it, sends it to the back, locks movement and selection, and then saves the file as an XLSX document while preserving all shape attributes.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Add a WordArt shape that will act as a watermark.
                // -----------------------------------------------------------------
                // Parameters: style, text, topRow, top, leftColumn, left, height, width
                Shape watermark = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,   // preset style
                    "CONFIDENTIAL",                    // watermark text
                    5,  // topRow (row index where the shape starts)
                    0,  // top offset in pixels
                    2,  // leftColumn (column index where the shape starts)
                    0,  // left offset in pixels
                    200, // height in pixels
                    600  // width in pixels
                );

                // -----------------------------------------------------------------
                // 3. Configure the watermark appearance.
                // -----------------------------------------------------------------
                // Make the text semi‑transparent.
                watermark.Fill.Transparency = 0.7; // 0 = opaque, 1 = fully transparent

                // Set a light gray fill color.
                watermark.Fill.FillType = FillType.Solid;
                watermark.Fill.SolidFill.Color = System.Drawing.Color.LightGray;

                // Remove the outline line (if supported by the current version).
                // Some older versions of Aspose.Cells do not expose IsVisible on LineFormat.
                // In such cases the line can be hidden by setting its weight to 0.
                if (watermark.Line != null)
                {
                    // Attempt to hide the line using the available API.
                    watermark.Line.Weight = 0;
                }

                // Rotate the watermark to a typical diagonal angle.
                watermark.RotationAngle = -45; // degrees

                // Send the shape to the back so that cell data is visible over it.
                watermark.ZOrderPosition = 0; // 0 = backmost

                // Lock the shape so it cannot be moved when the sheet is protected.
                // Use ShapeLockType.Move to prevent moving the shape.
                watermark.SetLockedProperty(ShapeLockType.Move, true);
                watermark.SetLockedProperty(ShapeLockType.Selection, true);

                // Verify that the shape is indeed WordArt.
                Console.WriteLine("Is the created shape WordArt? " + watermark.IsWordArt);

                // -----------------------------------------------------------------
                // 4. Save the workbook as XLSX preserving all shape attributes.
                // -----------------------------------------------------------------
                string outFile = "WordArtWatermark.xlsx";

                // Ensure the output directory exists.
                string outDir = Path.GetDirectoryName(outFile);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                workbook.Save(outFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
