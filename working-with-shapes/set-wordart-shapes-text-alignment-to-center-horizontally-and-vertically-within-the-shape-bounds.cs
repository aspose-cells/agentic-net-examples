// Title: How to center WordArt text horizontally and vertically in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set WordArt text horizontal alignment to center | center vertical text in WordArt shape with Aspose.Cells .NET | C# example aligning WordArt text both horizontally and vertically in Excel | using TextHorizontalAlignment and TextVerticalAlignment properties in Aspose.Cells | center WordArt shape text in generated Excel file using Aspose.Cells
// Tags: Aspose.Cells WordArt alignment API | TextHorizontalAlignment property usage | TextVerticalAlignment property usage | C# shape text alignment in Excel | Excel WordArt formatting with Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a WordArt shape to the first worksheet, sets TextHorizontalAlignment and TextVerticalAlignment to Center, and saves the file as WordArtCentered.xlsx.
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

            // Use the first preset WordArt style (Style1)
            PresetWordArtStyle style = (PresetWordArtStyle)0;

            // Add a WordArt shape to the worksheet.
            // In recent Aspose.Cells versions AddWordArt returns a Shape object.
            Shape wordArt = sheet.Shapes.AddWordArt(
                style,
                "WordArt Text",
                0,   // upper left row
                0,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                100, // height (pixels)
                200  // width (pixels)
            );

            // Center the text horizontally and vertically
            wordArt.TextHorizontalAlignment = TextAlignmentType.Center;
            wordArt.TextVerticalAlignment = TextAlignmentType.Center;

            // Save the workbook
            string outputPath = "WordArtCentered.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
