// Title: How to duplicate a WordArt (TextEffect) shape, reposition it to a specific cell range, and modify its text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that finds the first WordArt (TextEffect) shape in a worksheet, clones it, places the clone at cell B5, and sets its text to "Updated WordArt Text" using Aspose.Cells. | Write a method that copies a TextEffect shape, preserves its font, size, color, and dimensions, inserts the copy at a given row and column, and updates the caption in an Aspose.Cells workbook. | Create a reusable function in C# that accepts a source WordArt shape, target cell address, and new text, then duplicates the shape with original formatting and saves the workbook.
// Common Searches: aspnet aspose.cells duplicate wordart shape and move to cell B5 | c# copy text effect shape in Excel and change its text using Aspose | how to clone a WordArt object and reposition it in an Aspose.Cells worksheet | programmatically update WordArt caption in a .xlsx file with Aspose.Cells for .NET
// Tags: duplicate wordart shape aspose.cells c# | move texteffect shape to cell aspose.cells | update wordart text aspose.cells | clone shape preserving formatting aspose.cells | add wordart shape by cell coordinates c#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads or creates an Excel workbook, locates the first WordArt (TextEffect) shape, captures its formatting, adds a new TextEffect shape at cell B5 with the same appearance, changes its text to "Updated WordArt Text", and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Find the first WordArt (TextEffect) shape.
            Shape originalWordArt = null;
            foreach (Shape shape in sheet.Shapes)
            {
                // WordArt shapes have a non‑null TextEffect property.
                if (shape.TextEffect != null)
                {
                    originalWordArt = shape;
                    break;
                }
            }

            if (originalWordArt == null)
            {
                Console.WriteLine("No WordArt shape found in the worksheet.");
                workbook.Save(outputPath);
                return;
            }

            // Capture formatting from the original WordArt.
            string originalFontName = originalWordArt.Font.Name;
            int originalFontSize = (int)originalWordArt.Font.Size;
            bool isBold = originalWordArt.Font.IsBold;
            bool isItalic = originalWordArt.Font.IsItalic;
            Color originalFontColor = originalWordArt.Font.Color;
            double originalWidth = originalWordArt.Width;
            double originalHeight = originalWordArt.Height;

            // Target cell for the duplicated WordArt (e.g., B5).
            int targetRow = 4;    // zero‑based index (5th row)
            int targetColumn = 1; // zero‑based index (column B)

            // Add a new WordArt shape using a common preset.
            Shape duplicatedWordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset
                "Placeholder Text",              // temporary text
                originalFontName,
                originalFontSize,
                isBold,
                isItalic,
                targetRow,
                targetColumn,
                targetRow,
                targetColumn,
                0,
                0);

            // Apply original formatting.
            duplicatedWordArt.Font.Color = originalFontColor;
            duplicatedWordArt.Width = (int)originalWidth;
            duplicatedWordArt.Height = (int)originalHeight;

            // Update the text as required.
            duplicatedWordArt.Text = "Updated WordArt Text";

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
