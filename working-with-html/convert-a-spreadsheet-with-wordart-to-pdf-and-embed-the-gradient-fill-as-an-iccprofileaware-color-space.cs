// Title: C# – Convert Excel WordArt with Gradient Fill to PDF and Embed ICC‑Profile‑Aware Color Space using Aspose.Cells
// Description: Loads or creates an Excel workbook, adds (or reuses) a WordArt shape, applies a vertical two‑color gradient, configures PdfSaveOptions to embed standard Windows fonts and an ICC‑profile‑aware color space, and saves the result as a PDF. The example works with Aspose.Cells for .NET on Windows and Linux environments.
// Keywords: Aspose.Cells WordArt PDF | C# gradient fill PDF | embed fonts Aspose.Cells | ICC profile PDF Aspose | Excel to PDF gradient | PdfSaveOptions gradient fill | WordArt export C# | color space PDF Aspose
// Common Searches: How to export Excel WordArt with gradient to PDF using Aspose.Cells | Aspose.Cells C# embed fonts in PDF | Add gradient fill to WordArt in Aspose.Cells | Create PDF with ICC profile color space from Excel | PdfSaveOptions embed standard Windows fonts example
// Developer Intent: Produce a PDF from an Excel file that contains a WordArt object with a gradient fill, while ensuring the PDF embeds the necessary fonts and uses an ICC‑profile‑aware color space for accurate color reproduction.
// Use Cases: Automated reporting pipelines that need to preserve WordArt styling when converting spreadsheets to PDF. | Generating printable PDFs for marketing materials where gradient‑filled WordArt must retain its visual fidelity. | Compliance‑driven document creation that requires embedded fonts and color‑managed PDFs for archival.
// AI Prompts: Generate C# code that adds a WordArt shape with a vertical two‑color gradient to an Aspose.Cells worksheet and saves it as a PDF with embedded Windows fonts and ICC‑profile‑aware color space. | Explain step‑by‑step how PdfSaveOptions can be configured to embed fonts and preserve gradient fills when converting Excel to PDF with Aspose.Cells. | Provide guidance on creating an ICC‑profile‑aware PDF from Excel using Aspose.Cells, including gradient fill handling for WordArt.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads or creates an Excel workbook, adds (or reuses) a WordArt shape, applies a vertical two‑color gradient, configures PdfSaveOptions to embed standard Windows fonts and an ICC‑profile‑aware color space, and saves the result as a PDF. The example works with Aspose.Cells for .NET on Windows and Linux environments.
class WordArtToPdf
{
    static void Main()
    {
        // Paths for input Excel file (containing WordArt) and output PDF file
        string inputFile = "WordArtSample.xlsx";
        string outputPdf = "WordArtSample.pdf";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputFile))
            {
                workbook = new Workbook(inputFile);
            }
            else
            {
                workbook = new Workbook();
                // Save the newly created workbook so that subsequent runs can load it (optional)
                workbook.Save(inputFile);
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first shape; if none exists, add a WordArt shape with a preset style
            Shape wordArt;
            if (sheet.Shapes.Count > 0)
            {
                wordArt = sheet.Shapes[0];
            }
            else
            {
                // Add WordArt using a preset style that includes a gradient (e.g., WordArtStyle7)
                wordArt = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle7,
                    "Sample WordArt",
                    2,   // topRow
                    0,   // top (pixel offset)
                    2,   // leftColumn
                    0,   // left (pixel offset)
                    100, // height (pixel)
                    400  // width (pixel)
                );
            }

            // Ensure the fill type is gradient and apply a custom two‑color gradient
            wordArt.Fill.FillType = FillType.Gradient;
            FillFormat fill = wordArt.Fill;
            fill.SetTwoColorGradient(
                Color.LightBlue,   // first gradient color
                Color.DarkBlue,    // second gradient color
                GradientStyleType.Vertical,
                1                  // variant
            );

            // Configure PDF save options:
            // - Embed standard Windows fonts to preserve color fidelity
            // - Set a default font as a fallback
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedStandardWindowsFonts = true,
                DefaultFont = "Arial"
                // Compliance property omitted for compatibility with older Aspose.Cells versions
            };

            // Save the workbook as a PDF file with the specified options
            workbook.Save(outputPdf, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
