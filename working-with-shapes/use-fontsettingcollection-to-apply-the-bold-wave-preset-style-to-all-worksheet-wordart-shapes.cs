using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        ShapeCollection shapes = worksheet.Shapes;

        // Add sample WordArt shapes to demonstrate the styling
        shapes.AddWordArt(PresetWordArtStyle.WordArtStyle2, "First WordArt", 2, 10, 2, 10, 100, 200);
        shapes.AddWordArt(PresetWordArtStyle.WordArtStyle3, "Second WordArt", 5, 10, 5, 10, 100, 200);

        // Apply the Bold Wave preset style to every WordArt shape in the worksheet
        foreach (Shape shape in shapes)
        {
            if (shape.IsWordArt)
            {
                // Use FontSettingCollection to set a base WordArt style
                FontSettingCollection fontSettings = shape.TextBody;
                fontSettings.SetWordArtStyle(PresetWordArtStyle.WordArtStyle1); // base style

                // Use TextEffectFormat to make the text bold and apply the Wave shape
                TextEffectFormat textEffect = shape.TextEffect;
                textEffect.FontBold = true;
                textEffect.PresetShape = MsoPresetTextEffectShape.Wave1;
            }
        }

        // Save the workbook showing the applied styles
        workbook.Save("BoldWaveWordArt.xlsx");
    }
}