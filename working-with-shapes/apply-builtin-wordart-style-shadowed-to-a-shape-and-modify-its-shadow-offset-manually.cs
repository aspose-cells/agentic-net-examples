// Title: Apply WordArtStyle1 Shadow and Adjust Shadow Offset with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a WordArt shape using the built‑in WordArtStyle1 preset, changes the shadow preset to OffsetBottom, sets a custom shadow distance, and saves the file as XLSX.
// Keywords: Aspose.Cells | C# | WordArt | ShadowEffect | PresetWordArtStyle | PresetShadowType | Excel shape shadow | custom shadow offset | WordArtStyle1 | Aspose.Cells API | Excel automation
// Common Searches: Aspose.Cells set WordArt shadow offset C# | How to change WordArt shadow preset in Aspose.Cells | Apply built‑in WordArt style with shadow using Aspose.Cells | Modify ShadowEffect distance for WordArt shape | C# code to add WordArt with custom shadow in Excel
// Developer Intent: Add a WordArt shape with a built‑in shadow style, then programmatically modify the shadow preset and offset distance.
// Use Cases: Design a report header where WordArt needs a precise shadow placement for visual impact. | Generate marketing flyers in Excel that require consistent shadow positioning across multiple WordArt elements. | Standardize the shadow distance of all WordArt shapes in a workbook during automated report generation.
// AI Prompts: Show C# code to add a WordArt shape with WordArtStyle1 and then change its ShadowEffect preset and distance using Aspose.Cells. | Provide a script that iterates through every WordArt shape in a workbook and sets the shadow distance to 30 points. | Explain the differences between PresetShadowType.OffsetBottom and other shadow presets when customizing WordArt in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a WordArt shape using the built‑in WordArtStyle1 preset, changes the shadow preset to OffsetBottom, sets a custom shadow distance, and saves the file as XLSX.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape using a preset style that already contains a shadow (WordArtStyle1)
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style with shadow
            "Shadowed WordArt",                // text
            5,  // top row index
            0,  // vertical offset in pixels
            5,  // left column index
            0,  // horizontal offset in pixels
            200, // height in pixels
            400  // width in pixels
        );

        // Verify the shape is a WordArt object
        if (wordArt.IsWordArt)
        {
            // Apply a specific shadow preset (optional, can change from default)
            wordArt.ShadowEffect.PresetType = PresetShadowType.OffsetBottom;

            // Manually adjust the shadow offset by setting the distance (in points)
            wordArt.ShadowEffect.Distance = 40; // example offset
        }

        // Save the workbook
        workbook.Save("WordArtShadowed.xlsx", SaveFormat.Xlsx);
    }
}
