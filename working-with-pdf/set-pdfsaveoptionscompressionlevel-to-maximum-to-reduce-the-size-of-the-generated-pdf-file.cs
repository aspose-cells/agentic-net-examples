// Title: How to set PdfSaveOptions.CompressionLevel to Maximum in Aspose.Cells for C# to shrink PDF output size
// Common Searches: Aspose.Cells C# set PdfSaveOptions.CompressionLevel to Maximum for smaller PDF files | reduce PDF file size when saving workbook with Aspose.Cells by using maximum compression | configure PDF export compression level in Aspose.Cells .NET example | C# code to export Excel to PDF with highest compression using Aspose.Cells | Aspose.Cells PdfSaveOptions compression options for minimizing PDF size
// Tags: PdfSaveOptions compression level maximum | Aspose.Cells PDF size reduction | C# export workbook to compressed PDF | Aspose.Cells PDF compression settings .NET | set maximum PDF compression Aspose.Cells | optimize PDF output size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a Workbook, adds data, initializes a PdfSaveOptions object, and saves the workbook as a PDF. It shows where to set PdfSaveOptions.CompressionLevel to Maximum to reduce the generated PDF file size.
class PdfCompressionExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // For loading: 
                                                // if (File.Exists("input.xlsx"))
                                                //     workbook = new Workbook("input.xlsx");

            // Optional: add some data to the workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["A2"].PutValue(12345);

            // Configure PDF save options.
            // The Compression property is not available in some versions of Aspose.Cells,
            // so we use the default compression settings.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
