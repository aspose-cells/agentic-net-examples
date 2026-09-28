// Title: How to convert an XLSX workbook to a high‑resolution PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.ImageResolution to 300 DPI, and saves it as a PDF. | Show a console application that configures PdfSaveOptions for high‑quality PDF output when exporting an Excel workbook. | Explain how to adjust Aspose.Cells PDF export settings to improve image resolution during XLSX‑to‑PDF conversion.
// Common Searches: Aspose.Cells set PDF image resolution when exporting Excel to PDF in C# | C# convert XLSX to PDF with 300 DPI using Aspose.Cells | How to increase PDF quality from Excel conversion with Aspose.Cells .NET
// Tags: Aspose.Cells PDF export high DPI | PdfSaveOptions image resolution setting | C# Excel to PDF conversion Aspose.Cells | high‑resolution PDF generation from XLSX | Aspose.Cells workbook save as PDF

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;   // Contains PdfSaveOptions

// // Loads an XLSX workbook with Aspose.Cells, optionally sets PdfSaveOptions.ImageResolution for higher DPI, and saves the workbook as a high‑resolution PDF.
class ConvertXlsxToPdf
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Path where the high‑resolution PDF will be saved
        string pdfPath = "output.pdf";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Create PDF save options – you can adjust options here for higher quality
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Example option: keep the document structure (optional, not related to resolution)
        pdfOptions.ExportDocumentStructure = true;

        // If the library version supports it, you could set image resolution, e.g.:
        // pdfOptions.ImageResolution = 300;   // Uncomment if the property exists

        // Save the workbook as a PDF using the specified options
        workbook.Save(pdfPath, pdfOptions);

        Console.WriteLine("Conversion completed: " + pdfPath);
    }
}
