// Title: Insert a diagonal semi‑transparent WordArt watermark into a new Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a workbook, add a WordArt shape with custom text, rotate it -45°, apply 50% fill transparency, hide the outline, lock the shape, and save the file as .xlsx. | Show how to configure a WordArt shape as a free‑floating watermark in Aspose.Cells, including rotation, transparency settings, outline removal, and shape locking.
// Common Searches: asp.net add diagonal WordArt watermark to Excel workbook using Aspose.Cells | c# Aspose.Cells set shape transparency and lock for watermark | how to create semi transparent WordArt watermark in Excel with Aspose.Cells | Aspose.Cells rotate text effect shape as watermark in .xlsx | save Excel file with WordArt watermark using Aspose.Cells C#
// Tags: Aspose.Cells add WordArt watermark | Aspose.Cells diagonal WordArt | Aspose.Cells shape fill transparency | Aspose.Cells shape outline removal | Aspose.Cells shape locking | create new workbook with watermark Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, inserts a WordArt shape containing the text "CONFIDENTIAL" on the first worksheet, rotates it -45° to form a diagonal watermark, sets the fill to 50% transparency, makes the outline fully transparent, places the shape as free‑floating behind cells, locks it to prevent editing, and saves the workbook as WatermarkedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape that will serve as a watermark
            // Note: AddTextEffect overload requires two extra integer parameters (set to 0)
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "CONFIDENTIAL",
                "Arial",
                72,
                false,
                false,
                0,
                0,
                500,
                200,
                0,
                0);

            // Rotate the WordArt to appear as a diagonal watermark
            wordArt.RotationAngle = -45;

            // Make the WordArt semi‑transparent
            wordArt.Fill.Transparency = 0.5; // 50% transparent

            // Remove the outline by making it fully transparent
            wordArt.Line.Transparency = 1.0; // fully transparent outline

            // Send the shape behind the cells so it behaves like a watermark
            wordArt.Placement = PlacementType.FreeFloating;

            // Optionally lock the shape to prevent accidental editing
            wordArt.IsLocked = true;

            // Save the workbook to a file
            workbook.Save("WatermarkedWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
