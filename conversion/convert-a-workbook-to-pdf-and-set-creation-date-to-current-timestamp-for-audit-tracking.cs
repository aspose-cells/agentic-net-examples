// Title: Convert Excel to PDF with Current Creation Timestamp using Aspose.Cells for .NET
// Description: Loads an .xlsx workbook with Aspose.Cells, assigns DateTime.Now to PdfSaveOptions.CreatedTime, and saves the file as a PDF so the generated document contains a timestamp for audit tracking.
// Keywords: Aspose.Cells | C# PDF conversion | Excel to PDF | PdfSaveOptions | CreatedTime | PDF metadata timestamp | audit‑ready PDF | .NET Excel export | timestamped PDF
// Common Searches: Aspose.Cells set PDF creation date C# | PdfSaveOptions CreatedTime example | Add timestamp to PDF generated from Excel | Convert .xlsx to PDF with metadata using Aspose.Cells | How to embed audit timestamp in PDF with .NET
// Developer Intent: The developer needs to convert an Excel workbook to PDF and embed the current date‑time as the PDF's creation timestamp for audit purposes.
// Use Cases: Produce compliance‑focused PDF reports from financial spreadsheets that automatically record the export moment. | Schedule nightly conversion of dashboard workbooks to PDF with a built‑in audit trail. | Generate legally verifiable documents where the PDF metadata reflects the exact generation time.
// AI Prompts: Show how to add additional PDF metadata such as Author, Title, and Subject with PdfSaveOptions in Aspose.Cells. | Provide robust error handling for missing or inaccessible Excel files during timestamped PDF conversion. | Create a reusable C# method that takes an Excel path and returns a PDF byte array with CreatedTime set to UTC.

using System;
using Aspose.Cells;

namespace WorkbookToPdfWithTimestamp
{
    // Loads an .xlsx workbook with Aspose.Cells, assigns DateTime.Now to PdfSaveOptions.CreatedTime, and saves the file as a PDF so the generated document contains a timestamp for audit tracking.
    class Program
    {
        static void Main()
        {
            // Load the source Excel workbook (replace with actual file path)
            string sourcePath = "input.xlsx";
            Workbook workbook = new Workbook(sourcePath);

            // Create PDF save options and set the creation timestamp
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                CreatedTime = DateTime.Now // Audit tracking timestamp
            };

            // Save the workbook as PDF using the specified options
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook converted to PDF with CreatedTime = {pdfOptions.CreatedTime}");
        }
    }
}
