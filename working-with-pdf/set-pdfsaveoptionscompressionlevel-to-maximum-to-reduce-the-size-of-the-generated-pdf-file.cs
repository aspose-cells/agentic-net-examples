// Title: C# – Use Aspose.Cells PdfSaveOptions for Maximum PDF Compression (Minimum Size)
// Description: Shows how to create a workbook, configure PdfSaveOptions with Flate compression and MinimumSize optimization, and save it as the smallest possible PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | PdfSaveOptions | maximum compression | .NET | C# | Flate compression | PdfOptimizationType.MinimumSize | reduce PDF size | compact PDF | PDF compression
// Common Searches: Aspose.Cells set PDF compression to maximum | C# save workbook as small PDF | PdfSaveOptions MinimumSize example | How to reduce PDF size with Aspose.Cells .NET | Flate compression Aspose.Cells PDF
// Developer Intent: Generate a PDF from an Excel workbook with the highest compression settings to minimize file size.
// Use Cases: Email large spreadsheet reports as lightweight PDFs. | Archive financial statements where storage cost is a concern. | Deliver PDF invoices to customers on low‑bandwidth connections. | Embed PDFs in mobile apps where bundle size must stay minimal.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to PDF using PdfSaveOptions with maximum compression (Flate) and MinimumSize optimization. | Explain the differences between PdfCompressionCore.Flate and other compression options in Aspose.Cells. | What effect does PdfOptimizationType.MinimumSize have on image quality and file size in the generated PDF?

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to create a workbook, configure PdfSaveOptions with Flate compression and MinimumSize optimization, and save it as the smallest possible PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample PDF Compression Demo");
        worksheet.Cells["A2"].PutValue("This PDF is saved with maximum compression settings.");

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Set the core compression algorithm to Flate (the most efficient)
        pdfSaveOptions.PdfCompression = PdfCompressionCore.Flate;

        // Optimize for minimum file size (prioritizes size over print quality)
        pdfSaveOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF with the specified compression settings
        workbook.Save("output_max_compression.pdf", pdfSaveOptions);
    }
}
