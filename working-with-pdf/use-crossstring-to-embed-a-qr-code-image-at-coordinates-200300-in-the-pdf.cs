// Title: Create a QR code placeholder shape at absolute coordinates (200,300) and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// Common Searches: how to position a shape at exact point coordinates in a PDF with Aspose.Cells C# | Aspose.Cells C# place QR code placeholder at specific location before saving as PDF | absolute placement of free‑floating shapes in an Aspose.Cells workbook for PDF output | embed a shape at (200,300) points in a worksheet and generate a PDF with Aspose.Cells | C# Aspose.Cells add text effect shape with custom left and top values
// Tags: Aspose.Cells freefloating shape PDF | C# absolute shape placement Aspose.Cells | QR code placeholder shape Aspose.Cells | Save workbook as PDF with positioned shapes | AddTextEffect shape coordinates Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a free‑floating text effect shape as a QR code placeholder positioned at 200 points left and 300 points top, and saves the worksheet as a PDF file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a placeholder text shape where the QR code would appear.
            // The overload requires cell coordinates (leftColumn, topRow) followed by
            // position (left, top) and size (width, height) in points.
            Shape placeholder = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "QR Code Placeholder",
                "Arial",
                24,
                false,
                false,
                0,      // leftColumn (cell index)
                0,      // topRow (cell index)
                200,    // left (points)
                300,    // top (points)
                200,    // width (points)
                200);   // height (points)

            placeholder.Placement = PlacementType.FreeFloating; // allow absolute positioning

            // Define output file path
            string outputPath = "Output.pdf";

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
