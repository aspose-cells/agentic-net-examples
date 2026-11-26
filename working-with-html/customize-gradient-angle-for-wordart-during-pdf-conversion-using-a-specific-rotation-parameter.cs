// Title: Apply a custom gradient direction and rotate a WordArt shape before converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: how to set a two‑color gradient direction for WordArt using Aspose.Cells in C# | rotate a WordArt shape programmatically before saving an Excel file as PDF with Aspose.Cells | Aspose.Cells example for applying custom gradient and rotation to WordArt during PDF export | C# code to change WordArt rotation angle and gradient style when converting workbook to PDF | pdf conversion with rotated WordArt and horizontal gradient using Aspose.Cells .NET
// Tags: wordart two-color gradient aspose.cells c# | shape rotation angle wordart aspose.cells | pdf export wordart styling aspose.cells | horizontal gradient direction wordart aspose.cells | custom wordart appearance during pdf conversion

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace WordArtGradientPdfDemo
{
    // The sample creates a new workbook, inserts a WordArt shape, applies a two‑color horizontal gradient (LightBlue → DarkBlue), sets the shape's rotation angle to 30°, and then saves the workbook as a PDF using PdfSaveOptions with one page per sheet and PDF/A‑1b compliance.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPdfPath = "WordArtGradientDemo.pdf";

            try
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPdfPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a WordArt shape with a preset style
                Shape wordArt = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "Aspose.Cells WordArt",
                    2,   // top row index
                    10,  // top offset (pixels)
                    2,   // left column index
                    10,  // left offset (pixels)
                    100, // height (pixels)
                    400  // width (pixels)
                );

                // Apply a two‑color horizontal gradient (LightBlue → DarkBlue)
                // GradientStyleType.Horizontal corresponds to the desired direction
                // Variant value of 1 is a typical choice (range 1‑4)
                wordArt.Fill.SetTwoColorGradient(
                    Color.LightBlue,
                    Color.DarkBlue,
                    GradientStyleType.Horizontal,
                    1);

                // Rotate the whole WordArt shape (degrees)
                const float shapeRotation = 30f;
                wordArt.RotationAngle = shapeRotation;

                // Prepare PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true,
                    Compliance = PdfCompliance.PdfA1b
                };

                // Save the workbook as PDF
                workbook.Save(outputPdfPath, pdfOptions);

                Console.WriteLine($"PDF successfully created at: {outputPdfPath}");
                Console.WriteLine($"Shape rotation set to {shapeRotation}°.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while generating the PDF:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
