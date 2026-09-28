// Title: Add a WordArt (TextEffect) shape to an Excel worksheet, flip it horizontally with rotation, and verify the flip using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that inserts a TextEffect (WordArt) shape into a worksheet and creates a horizontal mirror effect by setting the shape's RotationAngle to 180 degrees. | Show how to programmatically confirm that the WordArt shape has been mirrored by checking that its RotationAngle property equals 180 after the transformation.
// Common Searches: C# Aspose.Cells how to mirror a WordArt shape horizontally | set rotation angle 180 for TextEffect shape in Aspose.Cells | verify that a WordArt shape is flipped in an Excel file using Aspose.Cells .NET | example of adding WordArt and rotating it with Aspose.Cells for .NET
// Tags: add wordart shape aspose.cells | horizontal mirror rotation aspose.cells | validate shape rotation aspose.cells c# | text effect shape excel .net | simulate flip using rotation aspose.cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The program creates a new workbook, adds a WordArt (TextEffect) shape to the first worksheet, simulates a horizontal flip by setting RotationAngle to 180°, checks that the rotation value is applied, and saves the file as WordArtFlipH.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Parameters for the shape position and size
            int upperLeftRow = 5;
            int upperLeftColumn = 5;
            int top = 0;
            int left = 0;
            int width = 200;
            int height = 50;

            // Add a WordArt (TextEffect) shape using the correct API
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset effect
                "Aspose.Cells",                  // text
                "Arial",                         // font name
                36,                              // font size
                false,                           // bold
                false,                           // italic
                upperLeftRow,
                upperLeftColumn,
                top,
                left,
                width,
                height);

            // Simulate horizontal flip by rotating 180 degrees
            wordArt.RotationAngle = 180;

            // Verify that the rotation (flip) was applied
            if (Math.Abs(wordArt.RotationAngle - 180) < 0.001)
            {
                Console.WriteLine("WordArt shape is flipped horizontally (simulated by 180° rotation).");
            }
            else
            {
                Console.WriteLine("Flip horizontally simulation not applied.");
            }

            // Save the workbook
            workbook.Save("WordArtFlipH.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
