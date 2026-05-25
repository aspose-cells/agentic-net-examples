using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class SetPdfTitleFromExcel
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(excelPath);

        // Ensure the workbook knows its original file name (useful for external references)
        workbook.FileName = excelPath;

        // Derive a title from the workbook name (without extension)
        string title = Path.GetFileNameWithoutExtension(excelPath);

        // Set the built‑in document property "Title" to the derived title
        workbook.BuiltInDocumentProperties.Title = title;

        // Create PDF save options and enable the display of the document title in the PDF window
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DisplayDocTitle = true
        };

        // Save the workbook as a PDF file; the PDF will carry the title metadata set above
        string pdfPath = "output.pdf";
        workbook.Save(pdfPath, pdfOptions);
    }
}