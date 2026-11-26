// Title: Create a reusable C# method to add a WordArt watermark with custom text, style, rotation, transparency, and lock settings using Aspose.Cells
// Common Searches: how to insert a WordArt watermark in an Excel file with Aspose.Cells C# | Aspose.Cells set shape lock properties for WordArt watermark in .NET | C# code to add semi transparent rotated WordArt as Excel worksheet watermark | protect worksheet while keeping WordArt watermark unselectable using Aspose.Cells | customize WordArt style and transparency for Excel watermark with Aspose.Cells
// Tags: add WordArt watermark Aspose.Cells | set shape lock properties Aspose.Cells | rotate WordArt shape Excel | transparent WordArt fill Aspose.Cells | WordArt placement MoveAndSize Aspose.Cells | preset WordArt style Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtWatermark
{
    // Provides a reusable C# method that adds a WordArt shape as a watermark to a worksheet, applies a preset style, rotates it -45°, sets 80% transparency, hides the outline, uses MoveAndSize placement, and optionally locks adjustment handles and selection for protected sheets.
    class Program
    {
        /// <param name="worksheet">Target worksheet.</param>
        /// <param name="text">Watermark text.</param>
        /// <param name="style">Preset WordArt style.</param>
        /// <param name="lockAdjustHandles">If true, the shape's adjustment handles are locked.</param>
        /// <param name="lockSelection">If true, the shape cannot be selected when the sheet is protected.</param>
        public static void AddWordArtWatermark(Worksheet worksheet, string text, PresetWordArtStyle style, bool lockAdjustHandles, bool lockSelection)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (string.IsNullOrEmpty(text)) throw new ArgumentException("Watermark text cannot be null or empty.", nameof(text));

            // Add WordArt shape. Parameters: style, text, topRow, top, leftColumn, left, height, width
            Shape wordArt = worksheet.Shapes.AddWordArt(
                style,
                text,
                0,          // topRow
                0,          // top (pixel offset)
                0,          // leftColumn
                0,          // left (pixel offset)
                500,        // height (pixels)
                1500);      // width (pixels)

            // Rotate to typical watermark angle
            wordArt.RotationAngle = -45;

            // Make the shape semi‑transparent
            wordArt.Fill.Transparency = 0.8; // 80% transparent

            // Hide outline by setting line weight to zero (Line.IsVisible not available in this version)
            wordArt.Line.Weight = 0;

            // Ensure the shape moves with cells but stays behind content
            wordArt.Placement = PlacementType.MoveAndSize;

            // Lock the desired properties
            wordArt.SetLockedProperty(ShapeLockType.AdjustHandles, lockAdjustHandles);
            wordArt.SetLockedProperty(ShapeLockType.Selection, lockSelection);
        }

        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data so the watermark effect can be seen
                sheet.Cells["A1"].PutValue("Sample Data");
                sheet.Cells["B2"].PutValue(123);
                sheet.Cells["C3"].PutValue(DateTime.Now);

                // Add a WordArt watermark
                AddWordArtWatermark(
                    sheet,
                    "CONFIDENTIAL",
                    PresetWordArtStyle.WordArtStyle1,
                    lockAdjustHandles: true,
                    lockSelection: true);

                // Protect the worksheet to enforce the lock settings (optional)
                sheet.Protect(ProtectionType.All);

                // Save the workbook
                string outputPath = "WordArtWatermarkDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
