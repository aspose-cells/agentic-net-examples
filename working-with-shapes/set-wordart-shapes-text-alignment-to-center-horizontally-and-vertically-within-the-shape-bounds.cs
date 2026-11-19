// Title: How to center text horizontally and vertically in a WordArt shape using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set WordArt text horizontal alignment to center | center vertical text in WordArt shape with Aspose.Cells .NET | programmatically align WordArt text in Excel workbook using Aspose.Cells | C# example for TextHorizontalAlignment and TextVerticalAlignment on WordArt | Aspose.Cells add WordArt and center its text in Excel file
// Tags: Aspose.Cells WordArt alignment API | TextHorizontalAlignment property usage | TextVerticalAlignment property usage | add WordArt shape to worksheet Aspose.Cells | center text inside shape Excel .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, adding a WordArt shape, setting both TextHorizontalAlignment and TextVerticalAlignment to Center, and saving the file as an XLSX workbook using Aspose.Cells for .NET.
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

            // Add a WordArt shape to the worksheet using the correct overload
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,          // preset style
                "Aspose.Cells WordArt",                    // text
                5, 5,                                      // upper‑left row and column
                0, 0,                                      // top and left offsets (in points)
                300, 100);                                 // height and width (in points)

            // Center the text horizontally
            wordArt.TextHorizontalAlignment = TextAlignmentType.Center;

            // Center the text vertically
            wordArt.TextVerticalAlignment = TextAlignmentType.Center;

            // Save the workbook to a file
            string outputPath = "WordArtCenteredAlignment.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
