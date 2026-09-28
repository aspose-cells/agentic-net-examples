// Title: Create a WordArt (TextEffect) shape with Comic Sans MS font and 48‑point size in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Insert a TextEffect shape into a worksheet, then set its FontName to "Comic Sans MS" and FontSize to 48 points with Aspose.Cells. | Generate a WordArt shape in an Excel file and programmatically customize its font family and size using C# and Aspose.Cells.
// Common Searches: Aspose.Cells C# how to create WordArt shape with specific font | set custom font family for TextEffect shape in Aspose.Cells .NET | change font size of WordArt (TextEffect) in Excel using Aspose.Cells | example code for adding WordArt with Comic Sans MS in Aspose.Cells
// Tags: add TextEffect shape Aspose.Cells | set WordArt font family C# | configure WordArt font size Aspose.Cells | customize WordArt shape Excel .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a WordArt (TextEffect) shape to the first worksheet, changes its font to Comic Sans MS, sets the font size to 48 points, and saves the workbook as WordArtExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Add a WordArt shape (TextEffect) to the worksheet
        // Parameters: preset effect, text, initial font name, initial font size, bold, italic,
        // upper left row, upper left column, top, left, width, height
        var shape = sheet.Shapes.AddTextEffect(
            MsoPresetTextEffect.TextEffect1,
            "Sample WordArt",
            "Calibri",          // initial font (will be changed)
            48,                 // initial size (will be set again)
            false,
            false,
            5,                  // row
            5,                  // column
            0,                  // top offset
            0,                  // left offset
            400,                // width
            100);               // height

        // Assign a custom font family
        shape.TextEffect.FontName = "Comic Sans MS";

        // Set the font size to forty‑eight points
        shape.TextEffect.FontSize = 48;

        // Save the workbook
        workbook.Save("WordArtExample.xlsx");
    }
}
