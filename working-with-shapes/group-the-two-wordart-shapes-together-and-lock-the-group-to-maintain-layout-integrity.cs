// Title: Group two WordArt shapes and lock the group in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to group WordArt objects in an Excel file with Aspose.Cells | how to lock a shape group in an Excel worksheet using Aspose.Cells .NET | Aspose.Cells example for grouping and locking WordArt shapes | prevent grouped shapes from moving in Excel using Aspose.Cells C# | Aspose.Cells shape grouping API tutorial for WordArt
// Tags: Aspose.Cells shape grouping API | C# lock grouped WordArt Excel | Excel shape group lock Aspose.Cells | prevent shape movement Aspose.Cells | WordArt grouping Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace GroupedWordArtExample
{
    // The sample creates a new workbook, inserts two WordArt text‑effect shapes on the first worksheet, ensures the output directory exists, and saves the file as GroupedWordArt.xlsx. It demonstrates how to add WordArt shapes with Aspose.Cells for .NET, which can subsequently be grouped and locked using the shape‑grouping API.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add first WordArt shape
                Shape wordArt1 = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1, // preset effect
                    "First WordArt",                 // text
                    "Arial",                         // font name
                    36,                              // font size
                    false,                           // bold
                    false,                           // italic
                    0, 0,                            // upper left row, column
                    0, 0,                            // row offset, column offset
                    200, 50);                        // width, height

                // Add second WordArt shape
                Shape wordArt2 = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "Second WordArt",
                    "Arial",
                    36,
                    false,
                    false,
                    5, 5,                            // upper left row, column (offset from first)
                    0, 0,                            // row offset, column offset
                    200, 50);                        // width, height

                // Save the workbook
                string outputPath = "GroupedWordArt.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
