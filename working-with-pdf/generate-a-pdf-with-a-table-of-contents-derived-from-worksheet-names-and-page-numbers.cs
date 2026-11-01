// Title: Create a multi‑page PDF from an Excel workbook where each worksheet starts on a new page using Aspose.Cells for .NET (C#)
// Common Searches: how to save each Excel worksheet as a separate page in a PDF with Aspose.Cells C# | Aspose.Cells PDFSaveOptions OnePagePerSheet example in .NET | C# convert workbook to PDF with one page per sheet using Aspose.Cells | generate PDF from Excel with page breaks between sheets in C# | Aspose.Cells export multiple worksheets to a single PDF file
// Tags: Aspose.Cells PDFSaveOptions OnePagePerSheet | export Excel worksheets to multi-page PDF C# | convert workbook to PDF with sheet-wise pagination | C# Aspose.Cells PDF generation per worksheet | save Excel as PDF with separate pages per sheet

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a Workbook, adds three worksheets filled with sample data, configures PdfSaveOptions.OnePagePerSheet to true, and saves the workbook as a PDF where each worksheet begins on its own page.
class Program
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a workbook and add sample worksheets
            // -------------------------------------------------
            var workbook = new Workbook();

            // Add three worksheets with some data
            for (int i = 0; i < 3; i++)
            {
                Worksheet ws = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"Sheet{i + 1}");

                // Simple header
                ws.Cells["A1"].PutValue("Header");

                // Fill 20 rows of sample data
                for (int row = 1; row <= 20; row++)
                {
                    ws.Cells[row, 0].PutValue($"Row {row} data for {ws.Name}");
                }
            }

            // -------------------------------------------------
            // 2. Save the workbook to PDF – each sheet on its own page
            // -------------------------------------------------
            var pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true // ensures each worksheet starts on a new page
            };

            string outputPdfPath = "WorkbookWithSheets.pdf";

            // Save the workbook as PDF
            workbook.Save(outputPdfPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPdfPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
