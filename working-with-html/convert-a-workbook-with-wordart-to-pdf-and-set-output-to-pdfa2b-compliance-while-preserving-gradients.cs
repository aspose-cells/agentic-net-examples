// Title: C# – Convert Workbook with Gradient WordArt to PDF/A‑2b Using Aspose.Cells
// Description: Demonstrates creating a workbook, adding a WordArt shape with a gradient fill, configuring PdfSaveOptions for PDF/A‑2b compliance, and saving the file as a PDF while preserving the gradient rendering.
// Keywords: Aspose.Cells C# PDF/A-2b | WordArt gradient export | PdfSaveOptions compliance | convert workbook to PDF/A-2b | preserve shape gradients Aspose | PDF/A-2b archival compliance | gradient WordArt PDF | Aspose.Cells PDF export | C# workbook to PDF/A-2b | PDF/A-2b compliance options
// Common Searches: C# Aspose.Cells export WordArt to PDF/A-2b | How to keep gradient fill when saving WordArt as PDF/A-2b | Set PDF/A-2b compliance in Aspose.Cells | Convert Excel with WordArt to archival PDF | PdfSaveOptions compliance property usage
// Developer Intent: Generate an archival‑grade PDF/A‑2b file from a spreadsheet that contains a gradient‑filled WordArt shape, ensuring the visual appearance remains unchanged.
// Use Cases: Produce marketing brochures with gradient WordArt and archive them as PDF/A‑2b for long‑term preservation. | Automate compliance‑ready PDF generation for financial reports that include styled WordArt elements. | Batch‑process multiple Excel files containing WordArt, converting each to PDF/A‑2b for legal documentation.
// AI Prompts: Write C# code with Aspose.Cells that adds a custom gradient WordArt shape and saves the workbook as PDF/A‑2b. | Explain the effect of setting PdfSaveOptions.Compliance = PdfCompliance.PdfA2b and list alternative compliance levels available in Aspose.Cells. | Provide a sample that converts an existing .xlsx file containing WordArt to PDF/A‑2b while preserving all shape formatting.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Demonstrates creating a workbook, adding a WordArt shape with a gradient fill, configuring PdfSaveOptions for PDF/A‑2b compliance, and saving the file as a PDF while preserving the gradient rendering.
class ConvertWordArtToPdfA2b
{
    static void Main()
    {
        // Create a new workbook (uses the Workbook constructor rule)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add WordArt with a gradient style (WordArtStyle7 = Gradient Fill - Blue, Accent 1, Reflection)
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            2, 0,   // row, top offset
            2, 0,   // column, left offset
            100,    // height
            400);   // width

        // Create PDF save options (uses PdfSaveOptions constructor rule)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set PDF/A‑2b compliance (uses PdfSaveOptions.Compliance property)
        pdfOptions.Compliance = PdfCompliance.PdfA2b;

        // Save the workbook as PDF with the specified compliance (uses Workbook.Save method)
        workbook.Save("WordArt_PdfA2b.pdf", pdfOptions);
    }
}
