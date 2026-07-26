// Title: Apply Wave WordArt style to a shape’s text using FontSettingCollection in Aspose.Cells for .NET (C#)
// Description: Create a workbook, add a WordArt shape, retrieve its FontSettingCollection, set WordArtStyle7, change the preset shape to Wave1, and save the file—all with Aspose.Cells C# API.
// Keywords: Aspose.Cells C# WordArt style | FontSettingCollection preset shape | Wave1 WordArt effect | WordArtStyle7 Aspose.Cells | programmatic WordArt styling .NET
// Common Searches: Aspose.Cells set Wave WordArt style C# | How to use FontSettingCollection for WordArt | Change WordArt preset shape to Wave1 in Excel | Apply WordArtStyle7 with Aspose.Cells | Add WordArt shape programmatically .NET
// Developer Intent: Programmatically apply the Wave preset WordArt style to the first shape’s text using FontSettingCollection.
// Use Cases: Design eye‑catching report headers with a wave‑styled title. | Automate dashboard visuals by applying consistent WordArt effects. | Generate branded Excel templates that include custom WordArt presets via code.
// AI Prompts: Generate C# code that adds a WordArt shape to an Excel sheet and sets its style to Wave using FontSettingCollection in Aspose.Cells. | Show how to change a WordArt shape’s preset shape to Wave1 after applying WordArtStyle7 with Aspose.Cells. | Explain the steps to retrieve a shape’s FontSettingCollection and modify both style and shape preset in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Create a workbook, add a WordArt shape, retrieve its FontSettingCollection, set WordArtStyle7, change the preset shape to Wave1, and save the file—all with Aspose.Cells C# API.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape (text effect) to the worksheet
        // Parameters: effect, text, font name, size, bold, italic, topRow, top, leftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddTextEffect(
            MsoPresetTextEffect.TextEffect1,   // basic text effect
            "Wave WordArt",                    // text to display
            "Arial",                           // font name
            36,                                // font size
            false,                             // not bold
            false,                             // not italic
            2,                                 // top row
            0,                                 // top offset (pixels)
            2,                                 // left column
            0,                                 // left offset (pixels)
            200,                               // height (pixels)
            400);                              // width (pixels)

        // Obtain the FontSettingCollection for the shape's text
        FontSettingCollection fontSettings = wordArt.TextBody;

        // Apply a preset WordArt style (example: WordArtStyle7)
        fontSettings.SetWordArtStyle(PresetWordArtStyle.WordArtStyle7);

        // Set the preset shape type to Wave1 to achieve the "Wave" appearance
        wordArt.TextEffect.PresetShape = MsoPresetTextEffectShape.Wave1;

        // Save the workbook to a file
        workbook.Save("WaveWordArtStyle.xlsx");
    }
}
