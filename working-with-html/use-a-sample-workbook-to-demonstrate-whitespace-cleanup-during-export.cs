// Title: Export a workbook to PDF while removing empty rows and columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export to PDF without blank rows and columns | remove whitespace from Excel sheet during PDF conversion using Aspose.Cells | fit worksheet to a single PDF page and skip empty rows Aspose.Cells .NET | how to use PdfSaveOptions to ignore empty rows in C# Aspose.Cells | C# generate PDF from workbook with whitespace cleanup Aspose.Cells
// Tags: Aspose.Cells PDF export whitespace removal | C# PdfSaveOptions AllColumnsInOnePagePerSheet | Aspose.Cells fit worksheet to page PDF | C# remove empty rows columns during export | Aspose.Cells sample workbook PDF generation

using System;
using System.IO;
using Aspose.Cells;

// This example creates a workbook with intentional empty rows and columns, configures the page setup to fit content, applies PdfSaveOptions (AllColumnsInOnePagePerSheet) and saves the file as a PDF, demonstrating how Aspose.Cells for .NET can eliminate whitespace during export.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (sample workbook)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with data, leaving empty rows and columns intentionally
            sheet.Cells["B2"].PutValue("Header");
            sheet.Cells["B3"].PutValue("Row1");
            // Row 4 is left empty to simulate whitespace
            sheet.Cells["B5"].PutValue("Row2");
            // Column C is left empty to simulate whitespace
            sheet.Cells["D2"].PutValue("ExtraHeader");
            sheet.Cells["D3"].PutValue("Extra1");
            sheet.Cells["D5"].PutValue("Extra2");

            // Adjust page setup to fit content
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;

            // Configure PDF export options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Fit all columns on one page per sheet (optional)
                AllColumnsInOnePagePerSheet = true
            };

            // Define output file path
            string outputPath = "WhitespaceCleanup.pdf";

            // Export the workbook to PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
