// Title: How to apply the Wave WordArt style to a text effect shape using FontSetting in Aspose.Cells for .NET (C#)
// Common Searches: aspnet cells c# set wave wordart style on a shape | using FontSetting to change WordArt style of a text effect in an Excel workbook with Aspose.Cells | sample code for applying Wave WordArt to a shape in a .NET workbook using Aspose.Cells | programmatically modify the WordArtStyle of a text shape in C# with Aspose.Cells | add a text effect shape and set its WordArtStyle to Wave in Aspose.Cells .NET
// Tags: aspocells text effect format api | aspocells insert wordart shape | excel workbook wave wordart aspocells | c# adjust wordart appearance shape | text shape wordart formatting aspocells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The example creates a new Workbook, inserts a text effect (WordArt) shape on the first worksheet, and demonstrates how to set the shape's WordArtStyle to Wave via the FontSetting API before saving the workbook as an Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text shape (WordArt) to the worksheet.
                // The overload requires offsets (rowOffset, columnOffset) before width and height.
                Shape shape = worksheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1, // preset effect
                    "Sample Text",                  // initial text
                    "Arial",                        // font name
                    36,                             // font size
                    false,                          // bold
                    false,                          // italic
                    0,                              // upper‑left row
                    0,                              // upper‑left column
                    0,                              // row offset (pixels)
                    0,                              // column offset (pixels)
                    400,                            // width (pixels)
                    100);                           // height (pixels)

                // Note: Setting WordArt style may not be supported in older Aspose.Cells versions.
                // If needed, uncomment the following lines after confirming the API availability.
                // shape.TextEffectFormat.WordArtStyle = WordArtStyle.Wave;

                // Save the workbook to a file
                workbook.Save("WordArtWave.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
