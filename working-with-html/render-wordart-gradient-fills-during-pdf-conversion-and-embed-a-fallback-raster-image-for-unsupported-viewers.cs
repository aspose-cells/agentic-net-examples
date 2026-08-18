// Title: Render Gradient WordArt in PDF with Raster Fallback Using Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds a WordArt shape with a two‑color horizontal gradient, inserts a PNG fallback picture aligned to the same size, sets the Z‑order so the WordArt overlays the raster image, and saves the sheet as a PDF with font‑compatibility options.
// Keywords: Aspose.Cells | WordArt gradient | PDF export | fallback image | AddPicture shape | Z‑order shapes | PdfSaveOptions | .NET Excel to PDF | gradient fill rendering
// Common Searches: Aspose.Cells render WordArt gradient in PDF | add fallback PNG for WordArt gradient Aspose.Cells | set shape Z‑order Aspose.Cells PDF | PdfSaveOptions font compatibility Aspose.Cells | gradient WordArt not showing in PDF viewer
// Developer Intent: Generate a PDF from an Excel workbook that preserves a gradient‑filled WordArt heading while providing a raster PNG fallback for PDF viewers that cannot render gradients.
// Use Cases: Create marketing brochures with decorative gradient WordArt titles that remain visible on legacy PDF readers. | Automate batch conversion of spreadsheets containing gradient WordArt, embedding aligned PNG fallbacks to avoid rendering issues. | Produce financial or technical reports where section headings use gradient WordArt, ensuring consistent appearance across all PDF viewers.
// AI Prompts: Write C# code with Aspose.Cells to add a WordArt shape using a two‑color horizontal gradient and embed a matching PNG fallback image before saving as PDF. | Explain how to control the Z‑order of shapes in Aspose.Cells so the WordArt appears above a raster fallback image in the exported PDF. | Show how to configure PdfSaveOptions for font compatibility and reliable gradient rendering when converting Excel to PDF with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// This C# example creates a workbook, adds a WordArt shape with a two‑color horizontal gradient, inserts a PNG fallback picture aligned to the same size, sets the Z‑order so the WordArt overlays the raster image, and saves the sheet as a PDF with font‑compatibility options.
class WordArtPdfWithFallback
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            ShapeCollection shapes = sheet.Shapes;

            // ------------------------------------------------------------
            // 1. Add a WordArt shape that uses a preset style with a gradient
            // ------------------------------------------------------------
            // WordArtStyle7 = Gradient Fill - Blue, Accent 1, Reflection
            Shape wordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7, // preset style with gradient
                "Gradient WordArt",               // text
                2,    // top row index
                10,   // vertical offset (pixels)
                2,    // left column index
                10,   // horizontal offset (pixels)
                100,  // height (pixels)
                400   // width (pixels)
            );

            // Ensure the fill type is gradient (optional, preset already sets it)
            wordArt.Fill.FillType = FillType.Gradient;

            // Apply a custom two‑color gradient (blue → light blue, horizontal)
            wordArt.Fill.SetTwoColorGradient(
                Color.Blue,          // first color
                Color.LightBlue,     // second color
                GradientStyleType.Horizontal,
                1                    // variant
            );

            // ------------------------------------------------------------
            // 2. Add a raster image as a fallback for viewers that cannot render the gradient
            // ------------------------------------------------------------
            string fallbackImagePath = "fallback.png"; // path to a raster image file

            if (File.Exists(fallbackImagePath))
            {
                try
                {
                    // Load the image into a stream because AddPicture expects a Stream
                    using (FileStream imgStream = new FileStream(fallbackImagePath, FileMode.Open, FileAccess.Read))
                    {
                        // Add the picture shape; cell range is arbitrary and will be adjusted later
                        Shape rasterFallback = shapes.AddPicture(
                            2, // upper left row
                            2, // upper left column
                            3, // lower right row
                            3, // lower right column
                            imgStream
                        );

                        // Align the raster picture with the WordArt dimensions.
                        rasterFallback.Top = wordArt.Top;
                        rasterFallback.Left = wordArt.Left;
                        rasterFallback.Height = wordArt.Height;
                        rasterFallback.Width = wordArt.Width;

                        // Send the raster shape to the back so the WordArt appears on top.
                        rasterFallback.ZOrderPosition = 0; // back
                        wordArt.ZOrderPosition = 1;        // front
                    }
                }
                catch (Exception imgEx)
                {
                    Console.WriteLine($"Failed to add fallback image: {imgEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Fallback image not found: {fallbackImagePath}");
            }

            // ------------------------------------------------------------
            // 3. Configure PDF save options and export the workbook
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure proper font handling for Unicode characters
                CheckWorkbookDefaultFont = true,
                CheckFontCompatibility = true
            };

            // Save the workbook as PDF.
            string outputPdf = "WordArtWithFallback.pdf";
            workbook.Save(outputPdf, pdfOptions);
            Console.WriteLine($"PDF generated: {outputPdf}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
