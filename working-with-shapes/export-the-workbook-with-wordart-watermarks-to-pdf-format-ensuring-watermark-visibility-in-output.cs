// Title: Create a diagonal WordArt watermark in an Excel sheet and export the workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds a WordArt text effect named 'CONFIDENTIAL' as a semi‑transparent diagonal watermark on the first worksheet, positions it behind the cells, and saves the workbook as a PDF with Aspose.Cells. | Show how to configure the Shape object's RotationAngle, Fill.Transparency, Placement, and ZOrderPosition properties so the watermark appears behind the worksheet content when converting to PDF. | Demonstrate setting the watermark width based on the worksheet's page size and exporting the workbook to PDF while preserving the watermark appearance.
// Common Searches: how to add a WordArt watermark to an Excel file with Aspose.Cells C# | Aspose.Cells C# diagonal semi transparent watermark behind cells | export Excel workbook to PDF preserving shape watermark using Aspose.Cells | set shape rotation and transparency in Aspose.Cells before PDF conversion | adjust watermark size based on page setup in Aspose.Cells C#
// Tags: wordart watermark Aspose.Cells | shape rotation Aspose.Cells | shape transparency Aspose.Cells | shape placement behind cells Aspose.Cells | pdf export with shape watermark Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a WordArt text effect labeled 'CONFIDENTIAL' as a diagonal semi‑transparent watermark behind the cells, rotates it -45°, sets its transparency to 0.5, adjusts its size based on page dimensions, and saves the workbook to PDF while preserving the watermark using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a default worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add WordArt (text effect) as a watermark.
            // Provide all required parameters, including height, shapeId and shapeType.
            Shape watermark = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "CONFIDENTIAL",
                "Arial",
                72,
                false,
                false,
                0,      // left
                0,      // top
                500,    // width
                100,    // height
                0,      // shapeId (default)
                0);     // shapeType (default)

            // Position and size the watermark
            watermark.Left = 0;
            watermark.Top = 0;
            watermark.Width = sheet.PageSetup.PaperSize == PaperSizeType.PaperA4 ? 595 : 612; // approximate page width in points
            watermark.Height = 100;

            // Rotate for diagonal appearance
            watermark.RotationAngle = -45;

            // Make semi‑transparent
            watermark.Fill.Transparency = 0.5; // 0 = opaque, 1 = fully transparent

            // Ensure the shape is printed and appears behind the cells
            watermark.Placement = PlacementType.FreeFloating;
            watermark.ZOrderPosition = 0; // send to back

            // Define output file path
            string outputPath = "WorkbookWithWatermark.pdf";

            // Save the workbook to PDF, preserving the watermark
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
