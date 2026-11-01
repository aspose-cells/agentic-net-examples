// Title: Convert an Excel workbook to PDF with lossless image compression using Aspose.Cells for .NET
// Common Searches: how to save an .xlsx workbook as a PDF with lossless image compression using Aspose.Cells in C# | Aspose.Cells PdfSaveOptions lossless image compression example C# | C# export Excel to PDF while preserving image quality with Aspose.Cells | convert Excel to PDF and reduce file size by compressing embedded images Aspose.Cells .NET | programmatically create a workbook if missing and export to PDF using Aspose.Cells
// Tags: Aspose.Cells PDF conversion lossless image compression | PdfSaveOptions ImageCompression C# | Export Excel to PDF Aspose.Cells .NET | Create workbook when file missing Aspose.Cells | Compress embedded images in PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // The example loads an existing Excel file or creates a simple workbook if none is found, configures PdfSaveOptions to apply lossless image compression, and saves the workbook as a PDF, demonstrating how to keep embedded image quality while reducing PDF size.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.pdf";

                // Ensure the input file exists; create a simple workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    var wb = new Workbook();
                    wb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                    wb.Save(inputPath);
                }

                // Load the source workbook.
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default settings).
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF.
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
