// Title: Add a diagonal semi‑transparent text watermark to each page when converting HTML to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook, inserts a WordArt watermark with custom text, font, size, rotation, and transparency on every worksheet, and saves the workbook as a PDF. | Show how to adjust the watermark properties—text, font name, font size, rotation angle, and transparency—in an Aspose.Cells HTML‑to‑PDF conversion example. | Provide a snippet that centers the watermark on an A4 page, sends it to the back layer, and ensures it appears behind cell content in the exported PDF.
// Common Searches: c# aspose.cells add diagonal watermark to pdf generated from html | how to set watermark transparency in aspose.cells html to pdf conversion | aspose.cells place text watermark behind cells when exporting workbook to pdf | change watermark rotation angle aspose.cells c# html to pdf | center watermark on each page aspose.cells pdf export
// Tags: Aspose.Cells HTML to PDF conversion with watermark | C# add diagonal text watermark Aspose.Cells | semi-transparent watermark shape Aspose.Cells | rotate text effect watermark worksheet | export workbook as PDF with watermark Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an HTML file into an Aspose.Cells Workbook, creates a WordArt shape containing a configurable text watermark, rotates it diagonally, centers it, applies 80% transparency, sends it to the back layer, and saves the workbook as a PDF so the watermark appears on every page.
class HtmlToPdfWithWatermark
{
    static void Main()
    {
        // Paths for input HTML and output PDF
        string htmlPath = "input.html";
        string pdfPath = "output.pdf";

        try
        {
            // Verify that the HTML file exists
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Input HTML file not found: {htmlPath}");

            // Load the HTML file into a Workbook using HtmlLoadOptions
            var loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Watermark settings
            string watermarkText = "CONFIDENTIAL";
            string fontName = "Arial";
            int fontSize = 72;               // Font size must be integer
            double rotationAngle = -45.0;    // Rotation in degrees

            // Add watermark to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add a WordArt shape (text effect) as watermark
                // Parameters: preset, text, font, size, bold, italic, left, top, width, height, depth, rotation
                Shape watermark = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    watermarkText,
                    fontName,
                    fontSize,
                    false,
                    false,
                    0,          // left
                    0,          // top
                    500,        // width
                    200,        // height
                    0,          // depth
                    0);         // rotation (handled later)

                // Rotate the shape diagonally
                watermark.RotationAngle = rotationAngle;

                // Make the watermark semi‑transparent
                watermark.Fill.Transparency = 0.8;          // 80% transparent fill
                watermark.Line.Transparency = 0.8;         // 80% transparent line

                // Send the shape to the back so it doesn't cover cell content
                watermark.ZOrderPosition = 0;

                // Approximate A4 page size in points (1 point = 1/72 inch)
                double pageWidth = 595;
                double pageHeight = 842;

                // Center the watermark on the page (cast to int if required by the API version)
                watermark.Left = (int)((pageWidth - watermark.Width) / 2);
                watermark.Top = (int)((pageHeight - watermark.Height) / 2);

                // Ensure the watermark stays fixed on the page
                watermark.Placement = PlacementType.FreeFloating;
            }

            // Save the workbook as PDF; the watermark will appear on each page
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF generated successfully: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
