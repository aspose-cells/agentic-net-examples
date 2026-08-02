// Title: Use LINQ to Filter WordArt Shapes and Apply Batch TextEffect Styling in Aspose.Cells (.NET)
// Description: Creates an Excel workbook, adds WordArt and regular shapes, then uses a LINQ query on the ShapeCollection to select only shapes where IsWordArt is true. For each WordArt shape the TextEffect format is updated (bold, italic, size, font, rotation, preset shape) and the workbook is saved.
// Keywords: Aspose.Cells | C# | LINQ | WordArt | ShapeCollection | IsWordArt | TextEffectFormat | batch styling | preset shape | Excel automation
// Common Searches: LINQ filter WordArt shapes Aspose.Cells | apply batch TextEffect formatting to WordArt in C# | select only WordArt from ShapeCollection | change preset shape of WordArt programmatically | update font properties for multiple WordArt objects
// Developer Intent: Select WordArt shapes from a worksheet and apply uniform TextEffect formatting in one pass.
// Use Cases: Standardize font style (bold, size, family) for all WordArt in a financial dashboard workbook. | Convert every WordArt to a specific preset shape (e.g., ArchUpCurve) for a marketing presentation. | Disable character rotation while applying consistent styling to WordArt in a template file.
// AI Prompts: Generate C# code that uses Aspose.Cells to retrieve only WordArt shapes with LINQ and set their font to Arial 18 bold. | Provide a LINQ query to filter ShapeCollection for WordArt and then apply a batch TextEffect change that sets FontItalic to false and PresetShape to StraightLine. | Write a method that iterates over all WordArt shapes in a workbook, disables RotatedChars, applies a preset shape, and saves the updated file.

using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtBatchStyle
{
    // Creates an Excel workbook, adds WordArt and regular shapes, then uses a LINQ query on the ShapeCollection to select only shapes where IsWordArt is true. For each WordArt shape the TextEffect format is updated (bold, italic, size, font, rotation, preset shape) and the workbook is saved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a few shapes: two WordArt and one regular rectangle
            Shape wordArt1 = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Hello World",
                2, 10,   // topRow, top
                2, 10,   // leftColumn, left
                100, 300 // height, width
            );

            Shape wordArt2 = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle5,
                "Aspose.Cells",
                5, 20,
                5, 20,
                120, 350
            );

            // Regular rectangle (not WordArt)
            worksheet.Shapes.AddRectangle(8, 30, 8, 30, 150, 200);

            // Use LINQ to filter only WordArt shapes
            var wordArtShapes = worksheet.Shapes
                                        .Where(shape => shape.IsWordArt)
                                        .ToList();

            // Apply batch style changes to each WordArt shape
            foreach (Shape shape in wordArtShapes)
            {
                // Access the TextEffect format of the WordArt shape
                TextEffectFormat textEffect = shape.TextEffect;

                // Example style changes
                textEffect.FontBold = true;          // Make text bold
                textEffect.FontItalic = true;        // Make text italic
                textEffect.FontSize = 24;            // Increase font size
                textEffect.FontName = "Calibri";     // Change font family
                textEffect.RotatedChars = false;    // Ensure characters are not rotated
                // Optionally change the preset shape type
                textEffect.PresetShape = MsoPresetTextEffectShape.ArchUpCurve;
            }

            // Save the workbook with the updated WordArt styles
            workbook.Save("WordArtBatchStyled.xlsx");
        }
    }
}
