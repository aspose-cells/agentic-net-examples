// Title: Export an Excel workbook containing hyperlinks to PDF and keep the links clickable with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a new workbook, adds a hyperlink to a cell, and saves it as a PDF using Aspose.Cells so the link stays clickable. | Extend the example to insert multiple hyperlinks across different rows and columns, then programmatically open the generated PDF to confirm each link is active. | Add comprehensive error handling that checks the PDF file existence, validates the number of hyperlinks, and logs success or failure messages.
// Common Searches: Aspose.Cells C# preserve Excel cell hyperlinks when converting to PDF | how to export workbook to PDF with active links using Aspose.Cells for .NET | C# generate PDF from Excel and retain clickable URLs | verify hyperlinks in PDF produced by Aspose.Cells conversion
// Tags: Aspose.Cells create hyperlink in Excel cell | export Excel to PDF with link retention | C# PDF export with hyperlink support | validate clickable URLs in generated PDF | hyperlink preservation during Excel to PDF conversion

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a workbook, inserting a hyperlink into a cell, exporting the workbook to PDF with Aspose.Cells, and confirming that the PDF file is generated successfully.
class HyperlinkPdfExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Write display text in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Visit Example.com");

            // Add a hyperlink to cell A1 (row 0, column 0)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
            sheet.Hyperlinks.Add(0, 0, 1, 1, "http://example.com");

            // Define PDF output path
            string pdfPath = "HyperlinkWorkbook.pdf";

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF file was created
            bool pdfCreated = File.Exists(pdfPath);
            Console.WriteLine(pdfCreated
                ? "PDF created successfully with hyperlink."
                : "Failed to create PDF.");
        }
        catch (Exception ex)
        {
            // Output any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
