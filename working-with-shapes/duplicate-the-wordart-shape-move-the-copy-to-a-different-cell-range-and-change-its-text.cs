// Title: Aspose.Cells C# – Duplicate a WordArt shape, move it to a new cell range, and edit its text
// Description: This example shows how to create a workbook, add a WordArt shape, clone it with the same size, reposition the copy using MoveToRange, modify its text via the TextEffect property, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | WordArt shape | duplicate shape | move shape to range | change WordArt text | AddWordArt | Shape.MoveToRange | TextEffect | Excel automation | programmatic Excel
// Common Searches: copy WordArt shape Aspose.Cells C# | move WordArt to specific cells programmatically | change text of WordArt using Aspose.Cells | duplicate shape and reposition in Excel worksheet | Aspose.Cells example for cloning WordArt
// Developer Intent: Create a copy of an existing WordArt object, place it in a different cell range, and update its displayed text.
// Use Cases: Generate section headers by duplicating a styled WordArt label across multiple worksheet areas. | Automate branding by copying a WordArt logo to several sheets and customizing the caption per sheet. | Build a template that repeats a WordArt tag for each page of a report, adjusting the text for each page.
// AI Prompts: Write C# code that clones a WordArt shape, moves it to rows 10‑12 and columns 3‑4, and sets its text to "Quarterly Summary" using Aspose.Cells. | Explain how Shape.MoveToRange positions a WordArt shape relative to cell boundaries in Aspose.Cells and which measurement units are applied.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtDuplicate
{
    // This example shows how to create a workbook, add a WordArt shape, clone it with the same size, reposition the copy using MoveToRange, modify its text via the TextEffect property, and save the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the shape collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add an original WordArt shape
            // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
            Shape originalWordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,   // preset style
                "Original WordArt",                 // initial text
                2, 0,                               // top row and vertical offset
                2, 0,                               // left column and horizontal offset
                100,                               // height in pixels
                300                                // width in pixels
            );

            // Duplicate the WordArt by adding a new shape with the same dimensions
            // Move the copy to a different cell range (e.g., rows 5-6, columns 5-6)
            Shape copyWordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,   // same preset style as original
                "Copy WordArt",                     // new text (will be set later)
                5, 0,                               // top row for the copy
                5, 0,                               // left column for the copy
                originalWordArt.Height,             // same height as original
                originalWordArt.Width               // same width as original
            );

            // Optionally, adjust the position more precisely using MoveToRange
            // MoveToRange(startRow, startColumn, endRow, endColumn)
            copyWordArt.MoveToRange(5, 5, 6, 6);

            // Change the text of the copied WordArt using TextEffect property
            if (copyWordArt.IsWordArt)
            {
                TextEffectFormat textEffect = copyWordArt.TextEffect;
                textEffect.Text = "Duplicated WordArt";
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("DuplicatedWordArt.xlsx");
        }
    }
}
