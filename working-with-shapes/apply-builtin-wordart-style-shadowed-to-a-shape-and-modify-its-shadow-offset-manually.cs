// Title: Add WordArt with Built‑in Shadow Style and Adjust Its Offset in Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a WordArt shape using the preset WordArtStyle1 (which includes a shadow), verifies the shape type, changes the shadow preset to OffsetBottom, manually sets the shadow distance to 40 points, and saves the file as an XLSX workbook.
// Keywords: Aspose.Cells WordArt shadow | C# WordArt shadow offset | preset WordArtStyle1 Aspose | ShadowEffect Distance property | modify WordArt shadow Aspose.Cells | shape shadow preset .NET
// Common Searches: Aspose.Cells change WordArt shadow distance | C# add WordArt with shadow in Excel file | set custom shadow offset for WordArt shape | how to modify ShadowEffect in Aspose.Cells | apply preset WordArt style with shadow using Aspose
// Developer Intent: Insert a WordArt shape with a built‑in shadow style and programmatically customize the shadow type and offset.
// Use Cases: Design a report header where the title uses WordArt with a bottom‑offset shadow for visual impact. | Generate marketing materials that require a specific WordArt style and precise shadow distance to match brand guidelines. | Automate the upgrade of existing spreadsheets by replacing plain text with WordArt that has a customized shadow effect.
// AI Prompts: Write C# code with Aspose.Cells to add a WordArt shape using the 'WordArtStyle1' preset and set its ShadowEffect to 'OffsetBottom' with a distance of 40 points. | Show how to retrieve an existing WordArt shape in a workbook and change its shadow preset type and distance before saving. | Explain the steps to access and modify the ShadowEffect properties of a WordArt shape in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a new workbook, inserts a WordArt shape using the preset WordArtStyle1 (which includes a shadow), verifies the shape type, changes the shadow preset to OffsetBottom, manually sets the shadow distance to 40 points, and saves the file as an XLSX workbook.
class ApplyWordArtShadow
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset style that already contains a shadow
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style with shadow
            "Shadowed WordArt",               // text to display
            5,   // top row index
            10,  // vertical offset (pixels)
            5,   // left column index
            10,  // horizontal offset (pixels)
            100, // height (pixels)
            400  // width (pixels)
        );

        // Verify the shape is WordArt before applying shadow settings
        if (wordArt.IsWordArt)
        {
            // Optionally set a specific preset shadow type
            wordArt.ShadowEffect.PresetType = PresetShadowType.OffsetBottom;

            // Manually adjust the shadow offset (distance) in points
            wordArt.ShadowEffect.Distance = 40; // custom offset
        }

        // Save the workbook with the WordArt shape
        workbook.Save("WordArtShadowed.xlsx", SaveFormat.Xlsx);
    }
}
