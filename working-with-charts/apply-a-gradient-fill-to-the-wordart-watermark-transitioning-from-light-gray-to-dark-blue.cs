// Title: Add a Light‑Gray to Dark‑Blue Horizontal Gradient WordArt Watermark in Aspose.Cells (C#)
// Description: Creates a new workbook, inserts a WordArt shape with the text "CONFIDENTIAL", applies a horizontal two‑color gradient (light gray → dark blue), sends the shape to the back as a watermark, and saves the file as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | gradient fill | WordArt watermark | horizontal gradient | two‑color gradient | shape fill type | PDF export | Excel workbook styling | Aspose.Cells shape
// Common Searches: Aspose.Cells add gradient WordArt watermark C# | horizontal two‑color gradient shape Aspose.Cells | export workbook with gradient watermark to PDF | set gradient fill on WordArt Aspose.Cells | C# code for gradient WordArt watermark in Excel
// Developer Intent: Programmatically add a WordArt watermark with a light‑gray to dark‑blue horizontal gradient and export the workbook as a PDF.
// Use Cases: Produce confidential PDF reports that display a subtle gradient watermark for branding or security. | Apply consistent gradient‑styled watermarks across multiple worksheets in automated report generation. | Enhance visual hierarchy in Excel dashboards by using gradient‑filled WordArt as background elements.
// AI Prompts: Generate C# code using Aspose.Cells to add a vertical red‑to‑orange gradient WordArt watermark and save as XLSX. | Show how to create a three‑color gradient fill on a shape in Aspose.Cells and export the result to PDF. | Provide an example of sending a gradient‑filled WordArt shape to the back of a sheet to act as a watermark in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a new workbook, inserts a WordArt shape with the text "CONFIDENTIAL", applies a horizontal two‑color gradient (light gray → dark blue), sends the shape to the back as a watermark, and saves the file as a PDF using Aspose.Cells for .NET.
class GradientWordArtWatermark
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape that will serve as the watermark
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1,   // base style (will be overridden by gradient)
            "CONFIDENTIAL",                     // watermark text
            0, 0,                               // row and top offset
            0, 0,                               // column and left offset
            100, 600);                          // height and width

        // Ensure the shape uses gradient fill
        wordArt.Fill.FillType = FillType.Gradient;

        // Obtain the GradientFill object and set a two‑color gradient
        GradientFill gradient = wordArt.Fill.GradientFill;
        gradient.SetTwoColorGradient(
            Color.LightGray,    // start color (light gray)
            Color.DarkBlue,     // end color (dark blue)
            GradientStyleType.Horizontal, // gradient direction
            1);                 // variant (default)

        // Optionally, send the shape to the back so it behaves like a watermark
        wordArt.ZOrderPosition = 0; // send to back

        // Save the workbook as PDF to view the watermark effect (lifecycle save)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("GradientWordArtWatermark.pdf", pdfOptions);
    }
}
