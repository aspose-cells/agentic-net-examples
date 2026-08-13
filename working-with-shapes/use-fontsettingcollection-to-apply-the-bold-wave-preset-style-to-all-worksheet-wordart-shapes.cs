// Title: C# – Apply Bold Wave WordArt Style to All Worksheet Shapes Using Aspose.Cells FontSettingCollection
// Description: This C# example creates a workbook, adds WordArt shapes, and iterates the worksheet's ShapeCollection. For each WordArt shape it sets a base WordArt style via FontSettingCollection, makes the text bold, applies the Wave1 text effect, and saves the file.
// Keywords: Aspose.Cells | C# FontSettingCollection | WordArt style | Bold Wave text effect | PresetWordArtStyle | ShapeCollection | Aspose.Cells .NET | TextEffectFormat | Excel automation | programmatic WordArt styling
// Common Searches: Aspose.Cells apply WordArt style to all shapes | C# set bold wave effect on WordArt | FontSettingCollection WordArt example | How to use PresetWordArtStyle in Aspose.Cells | Iterate worksheet shapes Aspose.Cells C#
// Developer Intent: Apply a bold wave WordArt style to every WordArt shape in a worksheet.
// Use Cases: Standardize heading WordArt across generated reports | Programmatically create marketing slides with uniform bold wave WordArt | Batch‑update existing Excel files to enforce a consistent WordArt appearance | Automate workbook styling for dashboards that use WordArt labels
// AI Prompts: Modify the sample to use PresetWordArtStyle.WordArtStyle5 while keeping the bold wave effect. | Add code that changes the font size to 24 pt and sets the text color to blue for all WordArt shapes. | Show how to apply the style only to shapes whose name starts with "Title_". | Explain how to retrieve and modify FontSettingCollection properties for existing WordArt in a loaded workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// This C# example creates a workbook, adds WordArt shapes, and iterates the worksheet's ShapeCollection. For each WordArt shape it sets a base WordArt style via FontSettingCollection, makes the text bold, applies the Wave1 text effect, and saves the file.
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

                // Enhance with bold font and a wave text effect
                TextEffectFormat textEffect = shape.TextEffect;
                textEffect.FontBold = true;
                textEffect.PresetShape = MsoPresetTextEffectShape.Wave1;
            }
        }

        // Save the workbook with the applied styles
        workbook.Save("BoldWaveWordArt.xlsx");
    }
}
