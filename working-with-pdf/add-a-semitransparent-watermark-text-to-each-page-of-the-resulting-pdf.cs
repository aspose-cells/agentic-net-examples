// Title: Create a semi‑transparent rotated text watermark on every worksheet and export to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds a 45‑degree, 50% transparent "CONFIDENTIAL" text effect shape as a watermark to each worksheet in an Aspose.Cells workbook before saving it as a PDF. | Demonstrate how to iterate through all worksheets, insert a free‑floating text effect shape with custom font, rotation, and transparency, and generate a PDF using Aspose.Cells for .NET.
// Common Searches: asp.net add 45 degree watermark to PDF generated from Excel with Aspose.Cells | c# Aspose.Cells rotate text shape and set transparency for PDF export | how to use AddTextEffect to create watermark in Aspose.Cells workbook | set watermark opacity when converting Excel to PDF using Aspose.Cells | apply CONFIDENTIAL watermark to every page of PDF from Excel C#
// Tags: Aspose.Cells text effect watermark | C# rotate shape transparency | export workbook to PDF with watermark | semi transparent watermark Excel PDF | free floating shape placement Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a workbook, adds a centered 45‑degree, 50% transparent "CONFIDENTIAL" text effect shape as a watermark to each worksheet, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example content – can be removed if loading an existing workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Watermark settings
            const string watermarkText = "CONFIDENTIAL";

            // Apply watermark text to every worksheet using a text effect shape
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Add a text effect shape that will act as a watermark
                // Note: AddTextEffect requires 12 parameters; the last two are anchor and text effect type (set to 0 for defaults)
                Shape watermarkShape = ws.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    watermarkText,
                    "Arial",
                    72,
                    false,
                    false,
                    0,      // left
                    0,      // top
                    500,    // width
                    100,    // height
                    0,      // anchor (default)
                    0);     // text effect (default)

                // Position the shape roughly at the center of the sheet
                watermarkShape.Left = 0;
                watermarkShape.Top = 0;

                // Rotate and make the shape semi‑transparent
                watermarkShape.RotationAngle = 45;                     // Rotate the watermark
                watermarkShape.Fill.Transparency = 0.5;               // 0 = opaque, 1 = fully transparent

                // Ensure the shape does not interfere with cell editing
                watermarkShape.Placement = PlacementType.FreeFloating;
                watermarkShape.IsLocked = false;
            }

            // Save the workbook as PDF – the watermark will appear on each page
            workbook.Save("Result.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
