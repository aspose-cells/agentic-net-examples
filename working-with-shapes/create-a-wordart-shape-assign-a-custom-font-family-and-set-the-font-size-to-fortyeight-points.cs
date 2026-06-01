using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class CreateCustomWordArt
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add a WordArt shape using AddTextEffect.
        // Parameters:
        //   effect      - preset text effect (choose any, here TextEffect1)
        //   text        - the WordArt text
        //   fontName    - custom font family
        //   size        - font size in points
        //   fontBold    - not bold
        //   fontItalic  - not italic
        //   topRow      - upper‑left row index
        //   top         - vertical offset in pixels
        //   leftColumn  - upper‑left column index
        //   left        - horizontal offset in pixels
        //   height      - shape height in pixels
        //   width       - shape width in pixels
        Shape wordArt = shapes.AddTextEffect(
            MsoPresetTextEffect.TextEffect1,
            "Custom WordArt",
            "Comic Sans MS",   // custom font family
            48,                // font size (points)
            false,
            false,
            2, 0,              // topRow, top
            2, 0,              // leftColumn, left
            200,               // height
            400);              // width

        // Ensure the font properties are set (optional, reinforces the settings)
        wordArt.TextEffect.FontName = "Comic Sans MS";
        wordArt.TextEffect.FontSize = 48;

        // Save the workbook with the WordArt shape
        workbook.Save("CustomWordArt.xlsx");
    }
}