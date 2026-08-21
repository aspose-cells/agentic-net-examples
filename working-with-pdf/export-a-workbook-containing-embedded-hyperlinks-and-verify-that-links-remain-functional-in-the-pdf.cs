// Title: Export Excel with Hyperlinks to PDF and Verify Links using Aspose.Cells for .NET
// Description: Creates a workbook, inserts a clickable URL into cell A1, saves the sheet as a PDF with PdfSaveOptions, then reads the PDF content to confirm the original hyperlink string is present, proving link preservation during conversion.
// Keywords: Aspose.Cells PDF export | Excel hyperlink to PDF | C# verify PDF link | PdfSaveOptions hyperlink | Aspose.Cells hyperlink preservation | validate PDF URLs .NET
// Common Searches: keep Excel hyperlinks when converting to PDF with Aspose.Cells | C# code to check if a PDF contains a specific URL | PdfSaveOptions retain links in exported PDF | how to validate hyperlinks in PDF generated from Excel
// Developer Intent: Generate a PDF from an Excel workbook that contains hyperlinks and programmatically ensure the links remain active after conversion.
// Use Cases: Producing PDF reports from spreadsheets where users must click through to external sites. | Automated testing pipelines that confirm hyperlink integrity after batch conversion of workbooks. | Compliance audits that require proof of link preservation in exported PDF documents.
// AI Prompts: Write C# that adds multiple hyperlinks to different cells, exports to PDF with Aspose.Cells, and verifies each URL in the resulting file. | Explain how to use Aspose.Pdf to open a PDF created by Aspose.Cells and programmatically click a hyperlink for integration testing. | Provide a guide to extract all URLs from a PDF generated from Excel using Aspose.Cells and list them in a console application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfSaveOptions

// Creates a workbook, inserts a clickable URL into cell A1, saves the sheet as a PDF with PdfSaveOptions, then reads the PDF content to confirm the original hyperlink string is present, proving link preservation during conversion.
class HyperlinkPdfDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put display text into cell A1
        sheet.Cells["A1"].PutValue("Visit Aspose");

        // Add a hyperlink to cell A1 (firstRow, firstColumn, totalRows, totalColumns, address)
        sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.aspose.com");

        // Save the workbook as PDF using PdfSaveOptions (save rule)
        string pdfPath = "HyperlinkDemo.pdf";
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save(pdfPath, pdfOptions);

        // Verify that the hyperlink URL is present in the generated PDF file
        // (simple verification by searching the raw PDF content)
        string pdfContent = File.ReadAllText(pdfPath);
        bool hyperlinkExists = pdfContent.Contains("https://www.aspose.com");

        Console.WriteLine("Hyperlink present in PDF: " + hyperlinkExists);
    }
}
