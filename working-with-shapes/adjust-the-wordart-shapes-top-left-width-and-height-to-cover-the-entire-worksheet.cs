// Title: How to expand an Aspose.Cells WordArt shape to cover the full worksheet in C#
// Common Searches: C# Aspose.Cells set WordArt shape size to worksheet dimensions | Resize WordArt to fill entire Excel sheet using Aspose.Cells API | Calculate total worksheet width and height for shape positioning in Aspose.Cells | Aspose.Cells WordArt top left zero full sheet coverage example | Programmatically make WordArt span whole worksheet in .NET
// Tags: Aspose.Cells WordArt full worksheet sizing | C# calculate worksheet dimensions points Aspose | set shape top left width height Aspose.Cells | resize WordArt to match sheet size .NET | shape positioning Excel using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads a workbook, adds a WordArt shape, computes the worksheet's total width and height in points by iterating over columns and rows, sets the shape's Top and Left to zero, assigns its Width and Height to the calculated totals, and saves the workbook so the WordArt covers the entire sheet.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape (provide required width and height parameters)
            Shape wordArt;
            try
            {
                wordArt = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1, // any preset style
                    "WordArt Text",
                    0, 0,               // upper‑left cell
                    0, 0,               // lower‑right cell (same as upper‑left for positioning)
                    100, 30);           // initial width and height in points
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to add WordArt: {ex.Message}");
                return;
            }

            // Approximate total worksheet width in points (1 character ≈ 7 points)
            double totalWidthPt = 0;
            for (int col = 0; col <= sheet.Cells.MaxColumn; col++)
            {
                totalWidthPt += sheet.Cells.GetColumnWidth(col) * 7.0;
            }

            // Total worksheet height in points (row height is already in points)
            double totalHeightPt = 0;
            for (int row = 0; row <= sheet.Cells.MaxRow; row++)
            {
                totalHeightPt += sheet.Cells.GetRowHeight(row);
            }

            // Position the WordArt at the top‑left corner and size it to cover the whole sheet
            wordArt.Top = 0;
            wordArt.Left = 0;
            wordArt.Width = (int)totalWidthPt;
            wordArt.Height = (int)totalHeightPt;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
