using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class WordArtShadowExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape using a preset style that already contains a shadow (WordArtStyle1)
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1,
            "Shadowed Text",
            2,   // topRow
            10,  // top (pixel offset)
            2,   // leftColumn
            10,  // left (pixel offset)
            200, // height (pixel)
            60   // width (pixel)
        );

        // Verify the shape is a WordArt object
        if (wordArt.IsWordArt)
        {
            // Access the shadow effect of the shape
            ShadowEffect shadow = wordArt.ShadowEffect;

            // Optionally change the preset shadow type (e.g., offset to the bottom)
            shadow.PresetType = PresetShadowType.OffsetBottom;

            // Manually adjust the shadow offset by setting the distance (in points)
            shadow.Distance = 30;          // offset distance
            shadow.Angle = 135;            // direction of the offset
            shadow.Blur = 20;              // blur amount
            shadow.Transparency = 0.4;     // 40% transparent
            // You can also set the shadow color if desired
            // shadow.Color = workbook.CreateCellsColor().Rgb = Color.FromArgb(128, 0, 0, 0);
        }

        // Save the workbook with the WordArt shape and customized shadow
        workbook.Save("WordArtShadowed.xlsx", SaveFormat.Xlsx);
    }
}