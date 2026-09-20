// Title: Add a WordArt TextEffect shape to the first worksheet and save the workbook as XLSX with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new Workbook, inserts a TextEffect WordArt shape with custom text, font, and size onto the first worksheet, and saves the file as an .xlsx using Aspose.Cells. | Explain how to modify curvature or other adjustment values of a WordArt shape in Aspose.Cells, and suggest alternative techniques when the AdjustmentValues property is unavailable.
// Common Searches: C# Aspose.Cells how to add WordArt TextEffect to a worksheet | Aspose.Cells TextEffect shape curvature adjustment not supported | save Excel file with WordArt using Aspose.Cells for .NET | example of creating WordArt shape in Aspose.Cells C# | workaround for WordArt adjustment values in Aspose.Cells
// Tags: add TextEffect WordArt shape Aspose.Cells C# | save workbook as XLSX with Aspose.Cells | WordArt curvature limitation Aspose.Cells | shape customization Aspose.Cells worksheet | C# Aspose.Cells TextEffect example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a TextEffect WordArt shape to the first worksheet, and saves the result as WordArtExample.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape (TextEffect) to the worksheet
            // Parameters: preset text effect, text, font name, font size, bold, italic,
            // left, top, width, height, anchor, text direction
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,   // preset effect
                "Aspose.Cells WordArt",            // displayed text
                "Arial",                           // font name
                48,                                // font size
                false,                             // bold
                false,                             // italic
                5,                                 // left position (points)
                5,                                 // top position (points)
                400,                               // width (points)
                100,                               // height (points)
                0,                                 // anchor (default)
                0);                                // text direction (default)

            // Note: AdjustmentValues property is not available in Aspose.Cells.
            // Curvature adjustments can be omitted or handled via other APIs if needed.

            // Save the workbook (lifecycle rule: save)
            workbook.Save("WordArtExample.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
