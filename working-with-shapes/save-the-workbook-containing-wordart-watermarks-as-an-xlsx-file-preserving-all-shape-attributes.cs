// Title: Add a WordArt Watermark, Preserve Shape Settings, and Save as XLSX with Aspose.Cells (.NET)
// Description: Shows how to create a Workbook, insert a WordArt shape as a diagonal semi‑transparent watermark, configure its rotation, transparency and Z‑order behind the cells, and save the file as an XLSX document while retaining all shape formatting.
// Keywords: Aspose.Cells WordArt watermark | C# Excel shape transparency | rotate WordArt Aspose.Cells | shape Z‑order Excel | save workbook with shapes | preserve shape properties .NET | add text effect shape | programmatic Excel watermark
// Common Searches: Aspose.Cells add WordArt watermark C# | how to rotate and make WordArt transparent in Excel using Aspose | save Excel file with shape Z‑order behind cells | preserve shape attributes when exporting to XLSX Aspose.Cells | batch apply watermark to multiple workbooks Aspose.Cells
// Developer Intent: Programmatically insert a WordArt watermark, adjust its visual attributes, and export the workbook to XLSX without losing any shape settings.
// Use Cases: Generate confidential reports with a diagonal, semi‑transparent watermark that stays behind all data. | Create corporate templates where branding WordArt must be embedded in every saved Excel file. | Automate the addition of the same watermark to a large set of workbooks while keeping exact shape formatting.
// AI Prompts: Write C# code using Aspose.Cells to add a WordArt watermark with custom text, font, size, rotation, and transparency, then save as XLSX. | Explain how to set the Z‑order of a shape so it appears behind worksheet cells in Aspose.Cells. | Provide a script that applies an identical WordArt watermark to all worksheets in an existing workbook, preserving shape properties.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a Workbook, insert a WordArt shape as a diagonal semi‑transparent watermark, configure its rotation, transparency and Z‑order behind the cells, and save the file as an XLSX document while retaining all shape formatting.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape that will serve as a watermark.
            // The AddTextEffect method requires two additional integer parameters (shape type and text effect).
            // Supplying default values (0) works for a basic watermark.
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

            // Rotate the shape to give a typical watermark appearance
            wordArt.RotationAngle = -45;

            // Make the shape semi‑transparent (80% transparent)
            wordArt.Fill.Transparency = 0.8;

            // Send the shape behind the cells (lower Z‑order)
            wordArt.ZOrderPosition = 0;

            // Save the workbook as XLSX
            workbook.Save("WatermarkWorkbook.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
