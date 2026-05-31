using System;
using Aspose.Cells;               // Core Aspose.Cells classes
using Aspose.Cells;               // For PdfSaveOptions (same namespace)

class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all hyperlinks from every worksheet to make the document static
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.Hyperlinks.Clear();
        }

        // Configure PDF save options (optional settings can be added here)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // ExportDocumentStructure can be set to true for better accessibility;
            // it does not affect hyperlink removal.
            ExportDocumentStructure = true
        };

        // Save the cleaned workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}