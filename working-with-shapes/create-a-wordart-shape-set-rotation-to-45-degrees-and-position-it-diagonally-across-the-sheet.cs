// Title: Create a 45-degree rotated WordArt (TextEffect) shape positioned diagonally across an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to add a TextEffect (WordArt) shape, set its placement to FreeFloating, define its height and width, rotate it 45 degrees, and span it from the top‑left to bottom‑right cells. | Show how to save the workbook after inserting and rotating the WordArt shape in Aspose.Cells. | Explain how to adjust the dimensions of a rotated TextEffect shape while preserving its rotation in Aspose.Cells.
// Common Searches: C# Aspose.Cells insert WordArt shape and rotate 45 degrees | how to set free‑floating placement for TextEffect shape in Aspose.Cells | position WordArt diagonally across an Excel sheet using Aspose.Cells | Aspose.Cells rotate shape programmatically example | adjust height and width of rotated WordArt in Aspose.Cells C#
// Tags: add TextEffect WordArt shape Aspose.Cells | rotate shape 45 degrees Aspose.Cells | free‑floating placement TextEffect Aspose.Cells | diagonal WordArt positioning Excel Aspose.Cells | set shape dimensions Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a TextEffect (WordArt) shape with custom size, sets it to free‑floating, rotates it 45 degrees, positions it diagonally across the sheet, and saves the file as WordArtDiagonal.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt (TextEffect) shape
            // Parameters: preset effect, text, font name, font size, bold, italic,
            // upper left row, upper left column, lower right row, lower right column,
            // height (in points), width (in points)
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset WordArt style
                "Aspose.Cells",                  // displayed text
                "Arial",                         // font name
                36,                              // font size
                false,                           // bold
                false,                           // italic
                0, 0,                            // start cell (top‑left)
                30, 15,                          // end cell (bottom‑right)
                200, 400);                       // height and width of the shape

            // Rotate the shape 45 degrees
            wordArt.RotationAngle = 45;

            // Ensure the shape is free‑floating so rotation works correctly
            wordArt.Placement = PlacementType.FreeFloating;

            // Save the workbook to a file
            string outputPath = "WordArtDiagonal.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
