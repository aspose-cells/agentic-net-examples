// Title: Lock a WordArt watermark in Excel using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts a WordArt watermark, locks selection, movement, resizing, handle adjustment, and text editing, then protects the worksheet so the watermark cannot be altered.
// Keywords: Aspose.Cells | C# | WordArt watermark | lock shape | ShapeLockType | worksheet protection | prevent editing | Excel security
// Common Searches: how to lock WordArt watermark Aspose.Cells | prevent moving WordArt shape in Excel C# | protect Excel worksheet objects with Aspose | disable editing of WordArt watermark .NET | shape lock properties Aspose.Cells example
// Developer Intent: Make a WordArt watermark immutable for end‑users.
// Use Cases: Add a confidential watermark to generated reports that cannot be changed. | Create a template where the logo WordArt stays fixed while cells stay editable. | Distribute spreadsheets with a protected branding element that recipients cannot modify.
// AI Prompts: Write C# code with Aspose.Cells to insert a WordArt watermark, lock all its properties, and protect the worksheet. | Explain the effect of each ShapeLockType option on a WordArt shape and how worksheet protection enforces them. | Show how to unlock a previously locked WordArt watermark for editing using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, inserts a WordArt watermark, locks selection, movement, resizing, handle adjustment, and text editing, then protects the worksheet so the watermark cannot be altered.
class LockWordArtWatermark
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape that will act as a watermark
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style
            "CONFIDENTIAL",                  // watermark text
            5, 5,                            // top row, top offset (pixels)
            5, 200,                          // left column, left offset (pixels)
            50, 300);                        // height, width (pixels)

        // Lock the shape itself
        wordArt.IsLocked = true;

        // Lock specific properties to prevent selection, moving, resizing, editing, etc.
        wordArt.SetLockedProperty(ShapeLockType.Selection, true);
        wordArt.SetLockedProperty(ShapeLockType.Move, true);
        wordArt.SetLockedProperty(ShapeLockType.Resize, true);
        wordArt.SetLockedProperty(ShapeLockType.AdjustHandles, true);
        wordArt.SetLockedProperty(ShapeLockType.Text, true);

        // Protect the worksheet so that the locked settings take effect
        worksheet.Protection.AllowEditingObject = false; // disallow editing of objects
        worksheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("LockedWatermark.xlsx");
    }
}
