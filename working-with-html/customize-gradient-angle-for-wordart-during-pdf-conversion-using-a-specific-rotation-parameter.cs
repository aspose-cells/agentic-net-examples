// Title: Set a Custom Gradient Angle for WordArt and Export to PDF with Aspose.Cells for .NET
// Description: Demonstrates how to add a WordArt shape, apply a linear gradient fill with a user‑defined angle, optionally rotate the text, and save the workbook as a PDF using Aspose.Cells in C#.
// Keywords: Aspose.Cells WordArt gradient angle | C# linear gradient fill WordArt | set GradientFill.Angle Aspose.Cells | WordArt text rotation Aspose.Cells | export workbook to PDF Aspose.Cells | custom gradient direction C# | shape FillType Gradient Aspose
// Common Searches: how to change gradient angle of WordArt in Aspose.Cells | set text rotation for WordArt shape C# | Aspose.Cells linear gradient fill angle example | export WordArt with custom gradient to PDF | C# Aspose.Cells gradient fill direction
// Developer Intent: Apply a specific gradient angle (and optional text rotation) to a WordArt shape before converting the worksheet to PDF.
// Use Cases: Design marketing brochures with angled WordArt headings that use a precise gradient direction. | Create report sections where the WordArt text is rotated for visual emphasis while maintaining consistent gradient shading. | Generate a series of worksheets, each containing WordArt with different gradient angles, and combine them into a single PDF document.
// AI Prompts: Write C# code using Aspose.Cells to add a WordArt shape, set a linear gradient fill with a custom angle, and save the workbook as a PDF. | Explain how to modify TextBody.TextAlignment.RotationAngle for a WordArt shape without affecting its gradient angle in Aspose.Cells. | Show how to change the GradientFill direction type (e.g., FromCenter, FromCorner) for WordArt and reflect the change in the exported PDF.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Demonstrates how to add a WordArt shape, apply a linear gradient fill with a user‑defined angle, optionally rotate the text, and save the workbook as a PDF using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset style that supports gradient fill
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,   // Gradient Fill - Blue, Accent 1, Reflection
            "Gradient WordArt",
            2,    // topRow
            0,    // top (pixels)
            2,    // leftColumn
            0,    // left (pixels)
            200,  // height (pixels)
            100   // width (pixels)
        );

        // Configure the shape to use a linear gradient fill
        wordArt.Fill.FillType = FillType.Gradient;
        // Initialize the gradient (angle will be overridden later)
        wordArt.Fill.GradientFill.SetGradient(GradientFillType.Linear, 0, GradientDirectionType.FromCenter);

        // Set the desired gradient angle (e.g., 60 degrees)
        wordArt.Fill.GradientFill.Angle = 60f;

        // Optional: rotate the text inside the WordArt shape (e.g., 45 degrees)
        // This uses the text alignment rotation property
        wordArt.TextBody.TextAlignment.RotationAngle = 45;

        // Save the workbook as PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("WordArtGradientAngle.pdf", pdfOptions);
    }
}
