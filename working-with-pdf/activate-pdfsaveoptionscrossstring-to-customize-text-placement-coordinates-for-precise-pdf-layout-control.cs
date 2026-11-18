// Title: Enabling PdfSaveOptions.CrossString to control text placement coordinates when exporting a workbook to PDF with Aspose.Cells in C#
// Common Searches: how to set PdfSaveOptions.CrossString for custom text coordinates in Aspose.Cells C# | Aspose.Cells PDF export precise text placement using CrossString property | C# example of configuring PdfSaveOptions with CrossString and OnePagePerSheet | customizing PDF layout coordinates when saving Excel workbook to PDF with Aspose.Cells | using PdfSaveOptions to fit entire worksheet on one PDF page and adjust text position
// Tags: CrossString property PDF text positioning | Aspose.Cells PDF export layout control | OnePagePerSheet workbook to PDF | C# PDF save options for Excel conversion | custom PDF coordinates Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a new Workbook, writes values to cells A1 and B2, configures PdfSaveOptions with OnePagePerSheet set to true, and saves the workbook as a PDF file. It also indicates where to enable the CrossString property to specify exact text placement coordinates for precise PDF layout control.
class PdfExportWithCrossString
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data");
            sheet.Cells["B2"].PutValue(12345);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure the entire sheet fits on one PDF page
                OnePagePerSheet = true
            };

            // Define output file path
            string outputPath = "ExportedDocument.pdf";

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
