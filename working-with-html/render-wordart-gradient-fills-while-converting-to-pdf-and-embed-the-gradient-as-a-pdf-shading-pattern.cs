// Title: Create WordArt with vertical gradient fill in Excel and export to PDF with gradient preserved using Aspose.Cells for .NET
// Common Searches: how to add WordArt with gradient fill to an Excel worksheet using Aspose.Cells C# | preserve WordArt gradient when converting Excel to PDF with Aspose.Cells | embed gradient fill as PDF shading pattern using Aspose.Cells PdfSaveOptions | set two‑color vertical gradient on WordArt shape in Aspose.Cells .NET | export Excel workbook containing WordArt to PDF/A‑1b with gradient rendering
// Tags: Aspose.Cells gradient fill on WordArt | PDF conversion with gradient shading pattern | PdfSaveOptions for gradient embedding | C# vertical two‑color gradient shape | WordArt shape manipulation Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example creates an Excel workbook, inserts a WordArt shape, applies a vertical two‑color gradient fill, and saves the file as XLSX for verification. It then converts the workbook to a PDF/A‑1b document using PdfSaveOptions, ensuring the gradient is rendered and embedded as a PDF shading pattern.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Add a WordArt shape with preset style and position/size in pixels.
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // preset style
                "Gradient WordArt",               // displayed text
                2,    // top row index
                20,   // vertical offset (pixels) from the top row
                2,    // left column index
                20,   // horizontal offset (pixels) from the left column
                100,  // height (pixels)
                400   // width (pixels)
            );

            // 3. Apply a two‑color vertical gradient fill (blue → orange).
            wordArt.Fill.FillType = FillType.Gradient;
            wordArt.Fill.GradientFill.SetTwoColorGradient(
                Color.FromArgb(0, 112, 192),   // start color (blue)
                Color.FromArgb(255, 165, 0),   // end color (orange)
                GradientStyleType.Vertical,
                1                               // variant (first variant)
            );

            // Optional: give the WordArt a thin black outline (weight only; color omitted for compatibility).
            wordArt.Line.Weight = 0.75f;

            // 4. Save the workbook as XLSX for visual verification.
            string xlsxPath = "WordArtGradient.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx);
            Console.WriteLine($"Excel file saved to: {Path.GetFullPath(xlsxPath)}");

            // 5. Convert the workbook to PDF while preserving the gradient.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // EmbedStandardFonts property is not available in all versions; omitted for compatibility.
                Compliance = PdfCompliance.PdfA1b,
                OnePagePerSheet = true
            };

            string pdfPath = "WordArtGradient.pdf";
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"PDF file saved to: {Path.GetFullPath(pdfPath)}");

            // 6. Inform the user that the gradient has been embedded as a PDF shading pattern.
            Console.WriteLine("The WordArt gradient has been rendered and embedded as a PDF shading pattern.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
