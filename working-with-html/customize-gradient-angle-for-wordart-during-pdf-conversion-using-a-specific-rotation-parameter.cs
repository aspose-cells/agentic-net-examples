// Title: Set WordArt Gradient Angle and Export to PDF with Aspose.Cells for .NET
// Description: Demonstrates how to add a WordArt shape, apply a linear gradient with a custom angle, and save the workbook as a PDF using Aspose.Cells. The example shows setting the gradient direction (e.g., 45°) and configuring two‑color gradient fills before conversion.
// Keywords: Aspose.Cells WordArt gradient angle | custom gradient direction Aspose.Cells | WordArt linear gradient .NET | export WordArt to PDF Aspose | gradient fill angle PDF conversion | C# Aspose.Cells shape styling
// Common Searches: how to change WordArt gradient angle in Aspose.Cells | Aspose.Cells set linear gradient direction for WordArt | customize WordArt gradient before PDF export .NET | Aspose.Cells gradient fill angle property example | C# WordArt gradient rotation Aspose.Cells
// Developer Intent: Apply a specific gradient angle to a WordArt shape so the angle appears correctly in the generated PDF.
// Use Cases: Design a report title with a 45° blue‑to‑light‑blue gradient and export it as a PDF. | Create marketing flyers where each WordArt element needs a distinct gradient direction before PDF conversion. | Allow end‑users to select gradient angles for WordArt in a web app and produce PDFs with the chosen visual style.
// AI Prompts: Show me C# code to set a custom gradient angle for a WordArt shape in Aspose.Cells and save it as a PDF. | Provide an Aspose.Cells example that changes the gradient direction of WordArt based on a variable angle and exports the workbook to PDF. | Explain how to combine shape rotation and gradient angle adjustments for WordArt using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsWordArtGradientPdf
{
    // Demonstrates how to add a WordArt shape, apply a linear gradient with a custom angle, and save the workbook as a PDF using Aspose.Cells. The example shows setting the gradient direction (e.g., 45°) and configuring two‑color gradient fills before conversion.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape with a preset style that supports gradient fill
            // Parameters: style, text, topRow, top, leftColumn, left, height, width
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
                "Gradient WordArt",
                2,   // topRow
                0,   // top (pixels)
                2,   // leftColumn
                0,   // left (pixels)
                100, // height (pixels)
                400  // width (pixels)
            );

            // Ensure the shape uses gradient fill
            wordArt.Fill.FillType = FillType.Gradient;

            // Set the gradient type to linear and define an initial angle (e.g., 0 degrees)
            wordArt.Fill.GradientFill.SetGradient(GradientFillType.Linear, 0f, GradientDirectionType.FromCenter);

            // Customize the gradient angle – this is the rotation parameter for the gradient
            // For example, set to 45 degrees
            wordArt.Fill.GradientFill.Angle = 45f;

            // Define the two colors for the gradient
            wordArt.Fill.GradientFill.SetTwoColorGradient(
                Color.Blue,          // Start color
                Color.LightBlue,     // End color
                GradientStyleType.Horizontal,
                1);

            // Optionally rotate the entire WordArt shape (not required for gradient angle)
            // wordArt.RotationAngle = 30; // Uncomment if shape rotation is desired

            // Prepare PDF save options (watermark not needed here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF; the gradient angle will be reflected in the output
            workbook.Save("WordArtGradient.pdf", pdfOptions);
        }
    }
}
