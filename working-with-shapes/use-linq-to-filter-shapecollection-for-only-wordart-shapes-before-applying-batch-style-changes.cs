using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtFilter
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the shape collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add a WordArt shape
            Shape wordArt1 = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Hello World",
                1,   // topRow
                10,  // top (pixel offset)
                1,   // leftColumn
                10,  // left (pixel offset)
                100, // height
                300  // width
            );

            // Add another WordArt shape
            Shape wordArt2 = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle5,
                "Aspose.Cells",
                5,
                20,
                5,
                20,
                120,
                350
            );

            // Add a regular rectangle shape (non‑WordArt)
            shapes.AddRectangle(10, 0, 10, 0, 150, 200);

            // Use LINQ to filter only WordArt shapes
            var wordArtShapes = shapes
                .Cast<Shape>()               // Convert ShapeCollection to IEnumerable<Shape>
                .Where(s => s.IsWordArt);    // Keep only shapes where IsWordArt is true

            // Apply batch style changes to each WordArt shape
            foreach (Shape shape in wordArtShapes)
            {
                // Access the TextEffect format of the WordArt shape
                TextEffectFormat textEffect = shape.TextEffect;

                // Example style changes
                textEffect.FontBold = true;
                textEffect.FontItalic = true;
                textEffect.FontName = "Calibri";
                textEffect.FontSize = 24;
            }

            // Save the workbook with the modified shapes
            workbook.Save("WordArtFilteredAndStyled.xlsx");
        }
    }
}