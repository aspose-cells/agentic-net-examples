using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;
using Aspose.Cells.Drawing; // for TextEffectFormat and enums

class ApplyWaveWordArtStyle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape (the first shape) with some sample text
        Shape shape = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 200, 50);
        shape.Text = "Wave WordArt Example";

        // Access the FontSettingCollection of the shape's text body
        FontSettingCollection fontSettings = shape.TextBody;

        // Apply a preset WordArt style (any style, here using WordArtStyle7 as an example)
        // This demonstrates the use of FontSettingCollection.SetWordArtStyle
        fontSettings.SetWordArtStyle(PresetWordArtStyle.WordArtStyle7);

        // Apply the Wave preset shape to the text effect of the shape
        // This gives the "Wave" visual effect to the WordArt text
        shape.TextEffect.PresetShape = MsoPresetTextEffectShape.Wave1;

        // Save the workbook to a file
        workbook.Save("WaveWordArtStyle.xlsx");
    }
}