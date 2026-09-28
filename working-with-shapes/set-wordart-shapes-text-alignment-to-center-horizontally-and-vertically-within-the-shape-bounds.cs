// Title: How to center text horizontally and vertically in a WordArt (TextEffect) shape using Aspose.Cells for .NET
// AI Prompts: Create a TextEffect (WordArt) shape on a worksheet and set its TextHorizontalAlignment and TextVerticalAlignment properties to TextAlignmentType.Center with Aspose.Cells in C#. | Adjust an existing WordArt shape so that its text is centered both horizontally and vertically within the shape bounds using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# set horizontal alignment for WordArt shape | center vertical text in TextEffect shape Aspose.Cells .NET | how to align WordArt text to middle of shape in Excel using Aspose.Cells | C# Aspose.Cells TextAlignmentType.Center for TextEffect shape
// Tags: Aspose.Cells TextEffect shape alignment | C# set WordArt text horizontal alignment | C# set WordArt text vertical alignment | Excel shape text centering with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example creates a new workbook, adds a WordArt (TextEffect) shape containing the text "Aspose", centers the text horizontally and vertically within the shape using TextAlignmentType.Center, and saves the file as WordArtAligned.xlsx.
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
            // upper left row, upper left column, height, width, lower right row, lower right column
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Aspose",
                "Arial",
                48,
                false,
                false,
                5,      // upper left row
                5,      // upper left column
                200,    // height
                100,    // width
                5,      // lower right row (using same as upper left for simplicity)
                5);     // lower right column

            // Center text horizontally within the shape
            wordArt.TextHorizontalAlignment = TextAlignmentType.Center;

            // Center text vertically within the shape
            wordArt.TextVerticalAlignment = TextAlignmentType.Center;

            // Save the workbook
            workbook.Save("WordArtAligned.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
