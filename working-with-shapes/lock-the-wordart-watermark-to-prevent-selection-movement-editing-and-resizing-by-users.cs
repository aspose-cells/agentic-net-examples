using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class LockWordArtWatermark
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a WordArt shape that will serve as the watermark
                // Parameters: style, text, topRow, top, leftColumn, left, height, width
                Shape wordArt = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "CONFIDENTIAL",
                    5,      // topRow
                    5,      // top (pixel offset)
                    5,      // leftColumn
                    5,      // left (pixel offset)
                    50,     // height (pixels)
                    300);   // width (pixels)

                // Lock the shape itself
                wordArt.IsLocked = true;

                // Lock specific properties to prevent user interaction when the sheet is protected
                wordArt.SetLockedProperty(ShapeLockType.Selection, true);   // cannot select
                wordArt.SetLockedProperty(ShapeLockType.Move, true);       // cannot move
                wordArt.SetLockedProperty(ShapeLockType.Resize, true);     // cannot resize
                wordArt.SetLockedProperty(ShapeLockType.Text, true);       // cannot edit text

                // Protect the worksheet (all protection types)
                sheet.Protect(ProtectionType.All);

                // Save the workbook
                string outputPath = "LockedWatermark.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            LockWordArtWatermark.Run();
        }
    }
}