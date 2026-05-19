using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtGroupDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add two WordArt (text effect) shapes
                // Parameters: preset effect, text, font name, font size, isBold, isItalic,
                // upperLeftRow, upperLeftColumn, top, left, width, height
                Shape wordArt1 = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "Hello", "Arial", 36, true, false,
                    2, 0, 0, 0, 200, 50);

                Shape wordArt2 = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "World", "Arial", 36, true, false,
                    5, 0, 0, 0, 200, 50);

                // Group the two WordArt shapes
                Shape[] shapesToGroup = new Shape[] { wordArt1, wordArt2 };
                GroupShape group = sheet.Shapes.Group(shapesToGroup);

                // Lock the group to prevent modifications when the sheet is protected
                group.IsLocked = true;

                // Save the workbook
                string outputPath = "WordArtGroupLocked.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}