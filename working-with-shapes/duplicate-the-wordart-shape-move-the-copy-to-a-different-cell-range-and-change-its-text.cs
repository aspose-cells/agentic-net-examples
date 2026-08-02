// Title: C# – Duplicate a WordArt Shape, Relocate It to a New Cell Range, and Edit Its Text Using Aspose.Cells
// Description: The sample creates a workbook, inserts a WordArt object, clones it, moves the copy to a different cell range with Shape.MoveToRange, updates the TextEffect content, and saves the result as WordArtDuplicate.xlsx. It illustrates the AddWordArt and MoveToRange methods of Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | WordArt | duplicate shape | move shape to cell range | Shape.MoveToRange | AddWordArt | Excel automation | change WordArt text
// Common Searches: How to copy a WordArt object and place it in another range with Aspose.Cells | Aspose.Cells C# move WordArt to specific rows and columns | Change the caption of a duplicated WordArt shape in .NET | Shape.MoveToRange example for Excel worksheets | AddWordArt and edit TextEffect using Aspose.Cells
// Developer Intent: Create a copy of an existing WordArt, reposition the copy in a different cell block, and assign new text to it.
// Use Cases: Generate a decorative header for a report, duplicate it, move the copy to a summary section, and give it a distinct label. | Place a branded WordArt logo in multiple worksheet zones, each with a customized caption, by cloning and relocating the shape. | Automate section headings across several sheets by reusing a base WordArt, shifting its location, and updating the displayed text.
// AI Prompts: Write C# code with Aspose.Cells that clones a WordArt shape, moves it to rows 10‑11 and columns 3‑4, and sets the text to "Quarterly Summary". | Explain how Shape.MoveToRange determines the position of a WordArt object within an Excel sheet when using Aspose.Cells for .NET. | Provide a step‑by‑step tutorial for copying a WordArt shape, modifying its TextEffect, and saving the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtDuplicate
{
    // The sample creates a workbook, inserts a WordArt object, clones it, moves the copy to a different cell range with Shape.MoveToRange, updates the TextEffect content, and saves the result as WordArtDuplicate.xlsx. It illustrates the AddWordArt and MoveToRange methods of Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the shape collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Define common parameters for the WordArt
            PresetWordArtStyle style = PresetWordArtStyle.WordArtStyle1;
            int topRow = 2;      // Upper left row index
            int top = 10;        // Vertical offset in pixels
            int leftColumn = 2;  // Upper left column index
            int left = 10;       // Horizontal offset in pixels
            int height = 50;     // Height in pixels
            int width = 200;     // Width in pixels

            // Add the original WordArt shape
            Shape originalWordArt = shapes.AddWordArt(
                style,
                "Original WordArt",
                topRow,
                top,
                leftColumn,
                left,
                height,
                width);

            // Add a duplicate WordArt with the same style and size
            Shape duplicateWordArt = shapes.AddWordArt(
                style,
                "Placeholder", // temporary text, will be changed later
                topRow,
                top,
                leftColumn,
                left,
                height,
                width);

            // Move the duplicate to a different cell range (e.g., rows 6-7, columns 6-7)
            // Parameters: topRow, leftColumn, bottomRow, rightColumn
            duplicateWordArt.MoveToRange(6, 6, 7, 7);

            // Change the text of the duplicate WordArt
            if (duplicateWordArt.IsWordArt)
            {
                TextEffectFormat textEffect = duplicateWordArt.TextEffect;
                textEffect.Text = "Duplicate WordArt";
            }

            // Save the workbook
            workbook.Save("WordArtDuplicate.xlsx");
        }
    }
}
