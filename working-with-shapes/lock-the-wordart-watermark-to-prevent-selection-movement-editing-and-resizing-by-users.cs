// Title: Lock a WordArt Watermark in Excel using Aspose.Cells for .NET (C#)
// Description: Shows how to insert a rotated WordArt watermark into an Excel workbook, lock it against selection, movement, resizing and text changes, and protect the worksheet so the watermark remains immutable, all with the Aspose.Cells C# API.
// Keywords: Aspose.Cells | C# | WordArt watermark | lock shape | ShapeLockType | worksheet protection | prevent editing Excel watermark | Excel shape lock | Aspose.Cells example | GitHub
// Common Searches: Aspose.Cells lock WordArt | prevent editing WordArt watermark in Excel C# | shape lock properties Aspose.Cells | protect worksheet objects Aspose.Cells | add rotated WordArt watermark Aspose.Cells
// Developer Intent: Prevent users from selecting, moving, resizing, or editing a WordArt watermark in an Excel file.
// Use Cases: Create confidential reports where the watermark cannot be altered by recipients. | Distribute Excel templates with a fixed branding watermark across all sheets. | Generate read‑only workbooks for regulatory compliance that keep the watermark intact.
// AI Prompts: Provide C# code to add a rotated WordArt watermark and lock it using Aspose.Cells. | Explain how to use ShapeLockType to disable selection, movement, resizing, and text editing of a shape in Aspose.Cells. | Show the steps to protect a worksheet while keeping a non‑editable WordArt watermark in an Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to insert a rotated WordArt watermark into an Excel workbook, lock it against selection, movement, resizing and text changes, and protect the worksheet so the watermark remains immutable, all with the Aspose.Cells C# API.
class WordArtWatermarkLockDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add WordArt as a watermark
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // preset style
                "CONFIDENTIAL",                  // watermark text
                0, 0,                            // top row, top offset (pixels)
                0, 0,                            // left column, left offset (pixels)
                200, 50);                        // height, width (pixels)

            // Position the WordArt (centered)
            wordArt.MoveToRange(0, 0, sheet.Cells.MaxDataRow, sheet.Cells.MaxDataColumn);

            // Rotate the WordArt for watermark effect
            wordArt.RotationAngle = -45;

            // Lock the shape to prevent selection, movement, editing, and resizing
            wordArt.IsLocked = true;
            wordArt.SetLockedProperty(ShapeLockType.Selection, true);
            wordArt.SetLockedProperty(ShapeLockType.Move, true);
            wordArt.SetLockedProperty(ShapeLockType.Resize, true);
            wordArt.SetLockedProperty(ShapeLockType.Text, true);

            // Protect the worksheet (disallow editing objects)
            sheet.Protection.AllowEditingObject = false;
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            string outputPath = "WordArtWatermarkLocked.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
