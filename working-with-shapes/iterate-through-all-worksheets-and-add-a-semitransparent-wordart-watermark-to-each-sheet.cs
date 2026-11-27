// Title: Add a semi‑transparent rotated WordArt watermark to every worksheet in an Excel file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add diagonal WordArt watermark to all sheets in a workbook | how to apply semi transparent watermark to each worksheet with Aspose.Cells .NET | iterate through worksheets and insert WordArt shape as watermark using Aspose.Cells | Aspose.Cells set WordArt shape transparency and rotation in Excel | create confidential watermark on multiple Excel sheets programmatically C#
// Tags: Aspose.Cells add WordArt shape to worksheet | WordArt shape transparency Aspose.Cells | rotate WordArt watermark Aspose.Cells | apply watermark to all worksheets Aspose.Cells | Excel workbook watermark C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WatermarkDemo
{
    // The sample loads an existing Excel workbook (or creates a new one), loops through each worksheet, and inserts a WordArt shape containing the text "CONFIDENTIAL". The shape is sized to cover the sheet, set to 50% transparency, rotated -45 degrees, and its outline is removed. After processing all sheets, the workbook is saved with the watermarks applied.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output_with_watermark.xlsx";

            Workbook workbook = null;

            try
            {
                if (File.Exists(inputPath))
                {
                    // Load existing workbook
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    // Create a new workbook if the file does not exist
                    workbook = new Workbook();
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Created a new workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            const string watermarkText = "CONFIDENTIAL";

            // Add watermark to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    ShapeCollection shapes = sheet.Shapes;

                    // Approximate worksheet size (pixels)
                    int sheetWidthPixels = sheet.Cells.MaxColumn * 64 + 200;
                    int sheetHeightPixels = sheet.Cells.MaxRow * 20 + 200;

                    // Add WordArt shape
                    Shape wordArt = shapes.AddWordArt(
                        PresetWordArtStyle.WordArtStyle1,
                        watermarkText,
                        0,               // topRow
                        0,               // top (pixels)
                        0,               // leftColumn
                        0,               // left (pixels)
                        sheetHeightPixels,
                        sheetWidthPixels);

                    // Make the text semi‑transparent
                    wordArt.Fill.Transparency = 0.5; // 0 = opaque, 1 = fully transparent

                    // Rotate the watermark diagonally
                    wordArt.RotationAngle = -45; // degrees

                    // Hide outline by setting line weight to zero
                    wordArt.Line.Weight = 0;

                    // Optional: set alternative text for accessibility
                    wordArt.AlternativeText = "Watermark";

                    Console.WriteLine($"Added watermark to sheet \"{sheet.Name}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add watermark to sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with watermarks to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
