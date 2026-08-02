// Title: C# – Convert Aspose.Cells Workbook to PDF and embed Author metadata
// Description: Shows how to assign the built‑in Author property of an Aspose.Cells workbook, configure PdfSaveOptions to export standard document properties, and save the workbook as a PDF that contains the author information in its metadata.
// Keywords: Aspose.Cells | C# | Workbook to PDF | Author property | Built‑in document properties | PdfSaveOptions | PDF metadata | Export document properties | Aspose.Cells PDF conversion
// Common Searches: Aspose.Cells set author before PDF export C# | How to include workbook author in PDF metadata using Aspose.Cells | PdfSaveOptions CustomPropertiesExport example | Export built‑in document properties to PDF with Aspose.Cells | C# code to add author to PDF generated from Excel
// Developer Intent: Convert an Excel workbook to PDF while preserving the Author metadata.
// Use Cases: Compliance reports that must display the creator’s name. | Automated invoice generation where the author identifies the issuing system or user. | Archival of financial statements with searchable author information. | Document management systems that index PDFs by author for quick retrieval.
// AI Prompts: Generate C# code to set Title and Subject properties and export them to PDF using Aspose.Cells. | Show how to read the Author field from a PDF created by Aspose.Cells. | Explain how to export custom document properties to PDF metadata with PdfSaveOptions. | Provide a step‑by‑step guide to batch convert multiple workbooks to PDF while assigning different authors.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to assign the built‑in Author property of an Aspose.Cells workbook, configure PdfSaveOptions to export standard document properties, and save the workbook as a PDF that contains the author information in its metadata.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the author property (built‑in document property)
        workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";

        // Add some sample data to demonstrate the workbook
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Workbook");
        sheet.Cells["A2"].PutValue("Created by: " + workbook.BuiltInDocumentProperties["Author"].Value);
        sheet.Cells["A3"].PutValue(DateTime.Now.ToString());

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Export built‑in/custom properties so the author appears in the PDF metadata
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}
