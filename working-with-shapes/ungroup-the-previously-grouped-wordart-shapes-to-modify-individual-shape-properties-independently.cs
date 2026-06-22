using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtUngroupDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two WordArt shapes to the worksheet
            Shape wordArt1 = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // preset style
                "Hello",                         // text
                1, 1,                            // upper left row, column
                200, 50,                         // height, width
                0, 0);                           // image width, image height (not used for WordArt)

            Shape wordArt2 = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle2,
                "World",
                3, 1,
                200, 50,
                0, 0);

            // Group the two WordArt shapes
            GroupShape group = worksheet.Shapes.Group(new Shape[] { wordArt1, wordArt2 });

            // Ungroup the shapes so they can be edited individually
            group.Ungroup();

            // Modify properties of each WordArt shape independently
            if (wordArt1.IsWordArt)
            {
                TextEffectFormat effect1 = wordArt1.TextEffect;
                effect1.FontSize = 24;      // larger font
                effect1.FontBold = true;    // bold text
                effect1.FontName = "Arial";
            }

            if (wordArt2.IsWordArt)
            {
                TextEffectFormat effect2 = wordArt2.TextEffect;
                effect2.FontSize = 18;      // smaller font
                effect2.FontItalic = true;  // italic text
                effect2.FontName = "Calibri";
            }

            // Save the workbook with the modified shapes
            workbook.Save("UngroupWordArtDemo.xlsx");
        }
    }
}