// Title: Convert Excel to PDF with timestamp using Aspose.Cells for .NET (C#)
// Description: Loads or creates an Excel workbook, sets PdfSaveOptions.CreatedTime to the current processing time, optionally applies PDF/A‑1b compliance, and saves the workbook as a PDF file with embedded creation‑date metadata.
// Keywords: Aspose.Cells | C# PDF conversion | Excel to PDF | PdfSaveOptions CreatedTime | PDF/A-1b compliance | set PDF creation date | timestamped PDF | Workbook.Save PDF | metadata timestamp
// Common Searches: Aspose.Cells set PDF creation date C# | Add timestamp to PDF generated from Excel | Save Excel as PDF/A-1b with Aspose.Cells | C# convert .xlsx to PDF with metadata | Aspose.Cells PDF metadata options
// Developer Intent: Generate a PDF from an Excel workbook and embed the current (or custom) processing time as the PDF's creation date, optionally enforcing PDF/A‑1b compliance.
// Use Cases: Produce audit‑ready reports where the PDF shows the exact generation time. | Create PDF/A‑1b compliant archival documents with a reliable timestamp. | Automate batch conversion of multiple workbooks to PDFs that share consistent metadata. | Supply legal or regulatory filings that require a documented creation date. | Integrate timestamped PDFs into downstream workflow systems for tracking.
// AI Prompts: Write C# code using Aspose.Cells to convert an .xlsx file to PDF and set the CreatedTime to now. | Show how to batch convert all Excel files in a directory to PDF/A‑1b with a processing timestamp using Aspose.Cells. | Explain how to read and modify PDF metadata after saving a workbook with Aspose.Cells. | Provide an example of assigning a custom creation date (e.g., a specific past date) in PdfSaveOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Needed for PdfCompliance enum

// Loads or creates an Excel workbook, sets PdfSaveOptions.CreatedTime to the current processing time, optionally applies PDF/A‑1b compliance, and saves the workbook as a PDF file with embedded creation‑date metadata.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx") to load

        // Add some data to demonstrate the conversion
        workbook.Worksheets[0].Cells["A1"].PutValue("Converted to PDF with processing timestamp");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set the PDF creation time to the current processing time
            CreatedTime = DateTime.Now,

            // Optional: set PDF compliance level (e.g., PDF/A-1b)
            Compliance = PdfCompliance.PdfA1b
        };

        // Define the output PDF file name
        string outputPdfPath = "ConvertedWorkbook.pdf";

        // Save the workbook as PDF using the configured options
        workbook.Save(outputPdfPath, pdfOptions);
    }
}
