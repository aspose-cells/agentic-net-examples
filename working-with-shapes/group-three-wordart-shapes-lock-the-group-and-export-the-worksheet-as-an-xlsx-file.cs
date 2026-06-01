using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        ShapeCollection shapes = worksheet.Shapes;

        // Add three WordArt shapes with different preset styles
        Shape wordArt1 = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // style
            "First",                         // text
            2, 10,                           // topRow, top (pixels)
            2, 10,                           // leftColumn, left (pixels)
            50, 200);                        // height, width (pixels)

        Shape wordArt2 = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle2,
            "Second",
            5, 10,
            5, 10,
            50, 200);

        Shape wordArt3 = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle3,
            "Third",
            8, 10,
            8, 10,
            50, 200);

        // Group the three WordArt shapes
        GroupShape group = shapes.Group(new Shape[] { wordArt1, wordArt2, wordArt3 });

        // Lock the group so it cannot be modified when the worksheet is protected
        group.IsLocked = true;

        // Save the workbook as an XLSX file
        workbook.Save("GroupedWordArt.xlsx");
    }
}