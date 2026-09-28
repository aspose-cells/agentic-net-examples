// Title: Add a 45‑degree semi‑transparent text watermark to odd‑numbered pages when exporting an Excel workbook to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that inserts a WordArt shape containing custom text, rotates it 45 degrees, sets its fill transparency to 60 % and sends it behind the worksheet content. | Adjust the workbook so the watermark shape is rendered only on odd‑numbered pages during PDF conversion with Aspose.Cells.
// Common Searches: how to add a semi‑transparent diagonal watermark to an Excel file using Aspose.Cells in C# | Aspose.Cells export to PDF with watermark only on odd numbered pages | set watermark opacity and angle with Aspose.Cells .NET API | create WordArt watermark for Excel worksheets before PDF conversion | C# example for applying custom text watermark on every other page in PDF output
// Tags: Aspose.Cells rotated watermark shape | C# shape transparency Aspose.Cells | WordArt text watermark Excel | odd pages PDF watermark Aspose.Cells | 45-degree watermark Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new Workbook, adds a WordArt shape with the text "CONFIDENTIAL" to the first worksheet, rotates the shape 45 degrees, sets its fill transparency to 60 %, hides the outline, sends the shape to the back, and saves the workbook. To limit the watermark to odd‑numbered pages, configure the PDF export options so the shape is applied only on those pages.
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

            // Add a WordArt shape that will act as a text watermark
            // Note: AddTextEffect returns a Shape object in recent Aspose.Cells versions
            Shape watermarkShape = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,   // preset effect
                "CONFIDENTIAL",                    // watermark text
                "Arial",                           // font name
                36,                                // font size
                false,                             // bold
                false,                             // italic
                0,                                 // upper-left column
                0,                                 // upper-left row
                500,                               // width (pixels)
                200,                               // height (pixels)
                0,                                 // upper-left column offset
                0);                                // upper-left row offset

            // Set rotation, transparency and hide outline
            watermarkShape.RotationAngle = 45;                     // 45‑degree rotation
            watermarkShape.Fill.Transparency = 0.6;               // 60% opacity (0 = opaque)
            watermarkShape.Line.Weight = 0;                       // hide border by setting line weight to zero

            // Send the shape to the back so it appears behind cell content
            watermarkShape.ZOrderPosition = 0;

            // Save the workbook
            string outputPath = "Watermarked.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
