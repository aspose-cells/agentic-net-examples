// Title: Create a PDF from an Aspose.Cells workbook with OnePagePerSheet enabled and a maximum of ten pages in C#
// AI Prompts: Write C# code that creates or loads an Aspose.Cells Workbook, sets PdfSaveOptions.OnePagePerSheet to true, sets PdfSaveOptions.MaxPages to 10, and saves the workbook as a PDF file. | Update an existing Aspose.Cells PDF conversion snippet to enforce a ten‑page limit while keeping each worksheet rendered on a single PDF page. | Describe how to configure PdfSaveOptions in Aspose.Cells for .NET to render each worksheet on one page and truncate the PDF after ten pages.
// Common Searches: Aspose.Cells C# PDF conversion one page per sheet limit to 10 pages | How to set MaxPages in PdfSaveOptions when saving Excel to PDF with Aspose.Cells | C# example for limiting PDF page count using Aspose.Cells PdfSaveOptions | OnePagePerSheet true and MaxPages 10 Aspose.Cells PDF output
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | Aspose.Cells limit PDF pages MaxPages | C# convert Excel to PDF with page count restriction | Aspose.Cells PDF rendering page limit | One page per worksheet PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, adds sample data, configures PdfSaveOptions with OnePagePerSheet set to true and MaxPages set to 10, then saves the workbook as a PDF while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate sample data (optional for PDF rendering)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B1"].PutValue("Data");
            sheet.Cells["B2"].PutValue(456);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true // Render each worksheet on a single PDF page
            };

            // Define output path
            string outputPath = "output.pdf";

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
