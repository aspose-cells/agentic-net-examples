// Title: Apply Wave WordArt Style to a TextBox Shape with FontSetting in Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds a TextBox shape, and uses FontSettingCollection.SetWordArtStyle (WordArtStyle7) together with TextEffectFormat.PresetShape = Wave1 to produce a wave WordArt effect, then saves the file.
// Keywords: Aspose.Cells C# WordArt | FontSetting SetWordArtStyle | Wave1 text effect | preset WordArt style Aspose.Cells | TextEffectFormat MsoPresetTextEffectShape | AddTextBox shape Excel | gradient WordArt Excel automation | Excel shape WordArt C#
// Common Searches: How to apply a preset WordArt style to a shape using Aspose.Cells C# | Set Wave text effect on a TextBox shape with FontSettingCollection | C# code for WordArtStyle7 in Aspose.Cells | Change shape to Wave1 preset in Aspose.Cells | Apply WordArt to Excel shapes programmatically
// Developer Intent: Use Aspose.Cells to add a TextBox shape and apply a wave WordArt style to its text via FontSettingCollection and TextEffectFormat.
// Use Cases: Design marketing dashboards where section titles appear with a blue gradient WordArt and wave shape for visual emphasis. | Automate report generation that highlights key metrics using a Wave1 preset WordArt effect on Excel text boxes. | Create printable Excel flyers where product names are displayed inside shapes with gradient WordArt styling.
// AI Prompts: Generate C# code that switches the WordArt style to WordArtStyle3 while keeping the Wave1 shape. | Show how to check if a shape supports TextEffect before applying the Wave1 preset in Aspose.Cells. | Explain how to customize the gradient colors of a WordArt style using FontSettingCollection.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;
using Aspose.Cells.Drawing; // for TextEffectFormat and related enums

// Create a new workbook and get the first worksheet
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];

// Add a text box shape (the first shape) to the worksheet
// Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 200, 100);

// Set some initial text for the shape
textBox.Text = "Wave WordArt Example";

// ------------------------------------------------------------
// Apply a preset WordArt style to the shape's text using FontSetting
// The TextBody property returns a FontSettingCollection which provides
// the SetWordArtStyle method.
// Here we use WordArtStyle7 as an example (gradient fill - blue, accent 1, reflection)
FontSettingCollection fontSettings = textBox.TextBody;
fontSettings.SetWordArtStyle(PresetWordArtStyle.WordArtStyle7);

// Additionally, set the preset shape type to a Wave (Wave1) to achieve the
// visual "wave" effect as requested.
if (textBox.IsWordArt) // Ensure the shape supports TextEffect
{
    TextEffectFormat textEffect = textBox.TextEffect;
    textEffect.PresetShape = MsoPresetTextEffectShape.Wave1;
}

// Save the workbook to a file
workbook.Save("WaveWordArtStyle.xlsx");
