// Title: Lock a WordArt watermark shape in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# lock WordArt shape so it cannot be selected or moved | prevent editing of Excel WordArt watermark using Aspose.Cells API | set shape lock properties for WordArt watermark in .NET workbook | protect worksheet to enforce locked WordArt watermark with Aspose.Cells | add and lock semi‑transparent diagonal WordArt watermark in C# Excel file
// Tags: Aspose.Cells shape lock API | Excel WordArt watermark protection C# | SetLockedProperty ShapeLockType usage | worksheet.Protect to enforce shape locks | add diagonal WordArt watermark Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtWatermarkLock
{
    // The example creates or loads an Excel workbook, inserts a WordArt shape as a semi‑transparent diagonal watermark, configures its fill and rotation, applies locked properties (selection, move, resize, adjust handles) via SetLockedProperty, protects the worksheet to enforce those locks, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if the file exists)
                Workbook workbook;
                string templatePath = "Template.xlsx";

                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a WordArt shape that will act as a watermark
                // Parameters: style, text, topRow, top, leftColumn, left, height, width
                Shape watermark = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "CONFIDENTIAL",
                    0,   // topRow
                    0,   // top (pixel offset)
                    0,   // leftColumn
                    0,   // left (pixel offset)
                    100, // height (pixels)
                    500  // width (pixels)
                );

                // Format the watermark (semi‑transparent, diagonal)
                watermark.Fill.FillType = FillType.Solid;
                watermark.Fill.SolidFill.Color = System.Drawing.Color.FromArgb(50, System.Drawing.Color.Gray);
                watermark.RotationAngle = -45;

                // Lock the shape so that users cannot select, move, resize or edit it
                watermark.SetLockedProperty(ShapeLockType.Selection, true);
                watermark.SetLockedProperty(ShapeLockType.Move, true);
                watermark.SetLockedProperty(ShapeLockType.Resize, true);
                watermark.SetLockedProperty(ShapeLockType.AdjustHandles, true);
                // Note: ShapeLockType.Delete is not available in the current API; omitted.

                // Protect the worksheet to enforce the locked properties
                worksheet.Protect(ProtectionType.All);

                // Save the workbook
                string outputPath = "WatermarkedLocked.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'. The WordArt watermark is locked.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
