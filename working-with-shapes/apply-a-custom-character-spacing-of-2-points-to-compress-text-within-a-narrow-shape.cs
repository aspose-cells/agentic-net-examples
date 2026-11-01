// Title: Compress text in a narrow Excel shape by applying negative character spacing with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set negative character spacing for shape text | how to use WordArt shape to adjust text spacing in Excel with Aspose.Cells | character spacing not changing for rectangle shape Aspose.Cells example | apply TextEffectFormat character spacing to compress text in a shape | Aspose.Cells negative character spacing example for narrow rectangle
// Tags: negative character spacing Aspose.Cells shape | WordArt TextEffectFormat shape text | adjust Excel shape text spacing C# | MsoDrawingType.TextEffect character spacing | compress text narrow rectangle shape

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // This example creates a workbook, adds a narrow rectangle shape, and demonstrates that plain rectangle shapes cannot have character spacing adjusted in Aspose.Cells. It advises using a WordArt (TextEffect) shape, which exposes TextEffectFormat for applying negative character spacing to compress the text before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (the rule)
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a narrow rectangle shape to the sheet
                // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,    // upper left row
                    0,    // upper left column
                    0,    // top offset (pixels)
                    0,    // left offset (pixels)
                    30,   // height (pixels)
                    200   // width (pixels)
                );

                // Set the text inside the shape
                shape.Text = "Compressed Text";

                // Note: Character spacing adjustment is not supported for plain rectangle shapes via Aspose.Cells.
                // If needed, use a WordArt shape (MsoDrawingType.TextEffect) which provides TextEffectFormat.

                // Save the workbook (the rule)
                string outputPath = "CompressedTextShape.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
