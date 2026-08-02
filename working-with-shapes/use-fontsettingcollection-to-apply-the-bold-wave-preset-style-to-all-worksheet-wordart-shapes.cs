// Title: Apply Bold Wave WordArt Style to All Shapes Using FontSettingCollection in Aspose.Cells for .NET
// Description: This C# example demonstrates how to load or create a workbook, locate every WordArt shape across all worksheets, retrieve its FontSettingCollection via TextBody, set a preset WordArt style, enable bold formatting, apply the Wave1 text‑effect shape, and save the result.
// Keywords: Aspose.Cells FontSettingCollection | WordArt style .NET | Bold Wave WordArt preset | apply TextEffect to shapes | iterate worksheets Aspose.Cells | C# Excel WordArt formatting | preset WordArt style programmatically
// Common Searches: how to set bold wave style for all WordArt in Aspose.Cells | FontSettingCollection example for WordArt | apply preset WordArt style to multiple shapes .NET | Aspose.Cells change TextEffect properties | C# code to format WordArt across worksheets
// Developer Intent: Programmatically apply a bold Wave1 preset WordArt style to every WordArt shape in all worksheets of an Excel workbook using Aspose.Cells.
// Use Cases: Standardize heading appearance in generated reports by applying a bold wave style to all WordArt titles. | Enforce corporate branding in Excel templates by automatically setting a specific WordArt preset on every shape. | Create visually consistent dashboards where all WordArt labels share the same bold wave effect.
// AI Prompts: Write C# code with Aspose.Cells that uses FontSettingCollection to make every WordArt shape bold and apply the Wave1 preset. | Show how to replace the preset WordArt style and the Wave1 shape with other PresetWordArtStyle and MsoPresetTextEffectShape values. | Provide a step‑by‑step tutorial for iterating worksheets and updating TextEffect properties of WordArt shapes in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;
using Aspose.Cells.Drawing;

// This C# example demonstrates how to load or create a workbook, locate every WordArt shape across all worksheets, retrieve its FontSettingCollection via TextBody, set a preset WordArt style, enable bold formatting, apply the Wave1 text‑effect shape, and save the result.
class ApplyBoldWaveToWordArt
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create

        // Add a sample worksheet with some WordArt shapes for demonstration
        Worksheet ws = workbook.Worksheets[0];
        ShapeCollection shapes = ws.Shapes;

        // Add a few WordArt shapes
        shapes.AddWordArt(PresetWordArtStyle.WordArtStyle1, "First", 2, 10, 2, 10, 100, 200);
        shapes.AddWordArt(PresetWordArtStyle.WordArtStyle2, "Second", 5, 10, 5, 10, 100, 200);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Process only WordArt shapes
                if (shape.IsWordArt)
                {
                    // Obtain the FontSettingCollection (TextBody) of the shape
                    FontSettingCollection fontSettings = shape.TextBody;

                    // Apply a preset WordArt style (choose any; here using WordArtStyle1)
                    fontSettings.SetWordArtStyle(PresetWordArtStyle.WordArtStyle1);

                    // Make the text bold
                    shape.TextEffect.FontBold = true;

                    // Set the wave preset shape
                    shape.TextEffect.PresetShape = MsoPresetTextEffectShape.Wave1;
                }
            }
        }

        // Save the workbook to a file
        workbook.Save("BoldWaveWordArt.xlsx"); // save
    }
}
