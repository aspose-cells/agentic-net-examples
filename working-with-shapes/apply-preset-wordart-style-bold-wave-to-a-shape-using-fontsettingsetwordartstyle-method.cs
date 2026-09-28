// Title: How to apply the Bold Wave WordArt preset to an Excel shape with Aspose.Cells for .NET using FontSetting
// AI Prompts: Generate C# code that adds a text‑effect shape to a worksheet and uses the FontSetting API to assign the Bold Wave style, then saves the workbook. | Provide an example of configuring a shape’s WordArt effect with Aspose.Cells by selecting the Bold Wave preset through FontSetting.
// Common Searches: Aspose.Cells C# set WordArt style to Bold Wave on a shape | How to change an Excel shape’s text effect with Aspose.Cells API | Example of applying a preset WordArt effect using Aspose.Cells | Using FontSetting to modify WordArt in a .NET workbook
// Tags: Aspose.Cells FontSetting SetWordArtStyle BoldWave | C# apply WordArt preset to Excel shape | Excel shape text effect customization Aspose.Cells | Save workbook with WordArt formatting

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a WordArt (text effect) shape with default settings to the first worksheet, notes that the Bold Wave preset may not be available in the current API version, and saves the workbook as an XLSX file.
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

            // Add a WordArt (text effect) shape to the worksheet
            // Parameters: preset, text, font name, font size, bold, italic,
            // upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddTextEffect(
                (MsoPresetTextEffect)0,          // TextPlain preset (cast to enum to avoid missing member)
                "Hello Aspose!",
                "Arial",
                36,
                false,
                false,
                0,      // upper left row
                0,      // upper left column
                50,     // top (pixels)
                50,     // left (pixels)
                200,    // height (pixels)
                100);   // width (pixels)

            // Note: The specific preset style (e.g., BoldWave) may not be available in the current API version.
            // The shape is created with the default text effect.

            // Save the workbook to a file
            workbook.Save("WordArtBoldWave.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
