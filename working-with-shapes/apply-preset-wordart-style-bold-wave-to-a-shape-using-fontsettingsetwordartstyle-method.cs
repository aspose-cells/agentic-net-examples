// Title: Apply Bold Wave WordArt to a TextBox shape with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts a TextBox, sets its text, applies PresetWordArtStyle7 via SetWordArtStyle, enables bold font, assigns the Wave1 text effect, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | WordArt style | SetWordArtStyle | Bold Wave | TextBox shape | PresetWordArtStyle7 | Wave1 text effect | Excel automation
// Common Searches: Aspose.Cells set WordArt style C# | How to add bold wave WordArt in Excel using Aspose | SetWordArtStyle example Aspose.Cells | Apply Wave1 text effect with Aspose.Cells | Create TextBox with WordArt in .NET
// Developer Intent: Insert a TextBox, apply a preset WordArt style, make the text bold, and set a wave effect using Aspose.Cells for .NET.
// Use Cases: Design a report header with eye‑catching bold wave WordArt for better visual hierarchy. | Generate marketing flyers where section titles need a wave‑shaped, bold WordArt appearance. | Automate Excel dashboards that require stylized headings without manual formatting.
// AI Prompts: Write C# code using Aspose.Cells to add a TextBox, apply PresetWordArtStyle7 with SetWordArtStyle, enable FontBold, and set PresetShape to Wave1. | Show how to change the WordArt style to PresetWordArtStyle5 and make the text italic in an Aspose.Cells workbook. | Explain how to detect if a shape is WordArt and modify its TextEffect properties (FontBold, PresetShape) in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;
using Aspose.Cells.Drawing;

// Create a new workbook and get the first worksheet
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];

// Add a TextBox shape that will hold the WordArt text
// Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 200, 200);

// Set the text that will receive the WordArt style
textBox.TextBody.Text = "Bold Wave Example";

// Apply a preset WordArt style to the text using FontSettingCollection.SetWordArtStyle
// Here we use WordArtStyle7 (Gradient Fill - Blue, Accent 1, Reflection) as an example
textBox.TextBody.SetWordArtStyle(PresetWordArtStyle.WordArtStyle7);

// Make the font bold and give the text a wave shape
if (textBox.IsWordArt)
{
    // FontBold property is part of TextEffectFormat
    textBox.TextEffect.FontBold = true;

    // Apply the "Wave1" preset shape to achieve a wave effect
    textBox.TextEffect.PresetShape = MsoPresetTextEffectShape.Wave1;
}

// Save the workbook to demonstrate the result
workbook.Save("BoldWaveWordArt.xlsx");
