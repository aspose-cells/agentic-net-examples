// Title: C# – Convert Aspose.Cells Workbook with Gradient WordArt to PDF/A‑2b
// Description: Shows how to insert a WordArt shape with a gradient (PresetWordArtStyle.WordArtStyle7) into an Aspose.Cells workbook and export it as a PDF/A‑2b compliant document using PdfSaveOptions, ensuring the gradient and other visual elements are retained.
// Keywords: Aspose.Cells PDF/A-2b | C# WordArt gradient | save workbook as PDF/A-2b | PdfSaveOptions compliance | .NET Excel to PDF/A-2b | preserve shapes Aspose.Cells | gradient WordArt export | PDF/A-2b conversion Aspose
// Common Searches: Aspose.Cells export WordArt to PDF/A-2b C# | preserve gradient WordArt when saving PDF/A-2b | PdfSaveOptions PdfCompliance.PdfA2b example | convert Excel with WordArt to PDF/A-2b | C# Aspose.Cells PDF/A-2b compliance guide
// Developer Intent: Generate a PDF/A‑2b file from an Excel workbook that contains gradient WordArt while keeping the original visual fidelity.
// Use Cases: Archiving design‑rich reports with gradient headings in a PDF/A‑2b format for long‑term preservation. | Automating the production of marketing brochures that include stylized WordArt and must meet PDF/A‑2b standards. | Batch‑processing multiple Excel files containing WordArt, converting each to PDF/A‑2b without losing gradients.
// AI Prompts: Write C# code that adds a gradient WordArt shape to an Aspose.Cells worksheet and saves the file as PDF/A‑2b. | Explain how PdfSaveOptions.Compliance = PdfCompliance.PdfA2b preserves gradients and what additional settings may be required. | Provide troubleshooting steps if the gradient appears flat after converting an Aspose.Cells workbook to PDF/A‑2b.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Shows how to insert a WordArt shape with a gradient (PresetWordArtStyle.WordArtStyle7) into an Aspose.Cells workbook and export it as a PDF/A‑2b compliant document using PdfSaveOptions, ensuring the gradient and other visual elements are retained.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a gradient style (WordArtStyle7)
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            2, 0,   // row, top offset
            2, 0,   // column, left offset
            100,    // height
            400);   // width

        // Set PDF save options to PDF/A‑2b compliance (preserves gradients)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2b
        };

        // Save the workbook as a PDF file with the specified compliance level
        workbook.Save("WordArt_PdfA2b.pdf", pdfOptions);
    }
}
