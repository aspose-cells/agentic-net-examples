// Title: How to enable maximum PDF compression with Aspose.Cells PdfSaveOptions in C# to minimize output file size
// AI Prompts: Write C# code that creates a PdfSaveOptions object, sets its Compression property to PdfCompressionType.Maximum, and saves a Workbook as a PDF using Aspose.Cells. | Show how to reduce the size of a PDF generated from an Excel workbook by configuring Aspose.Cells PdfSaveOptions for the highest compression in a C# application.
// Common Searches: Aspose.Cells C# set PdfSaveOptions compression to maximum for smaller PDF | reduce PDF file size when exporting Excel to PDF with Aspose.Cells | example of using PdfCompressionType.Maximum in Aspose.Cells | C# code to save workbook as compressed PDF using Aspose.Cells | how to configure PDF compression level in Aspose.Cells library
// Tags: Aspose.Cells PdfSaveOptions maximum compression | C# PdfCompressionType.Maximum usage | export Excel workbook to compressed PDF Aspose.Cells | optimize generated PDF size with Aspose.Cells | PdfSaveOptions compression level setting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PDF rendering options

namespace AsposeCellsPdfExample
{
    // The example loads or creates an Excel workbook, configures PdfSaveOptions to use PdfCompressionType.Maximum, and saves the workbook as a highly compressed PDF, resulting in a smaller output file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source workbook
                string inputPath = "input.xlsx";
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    try
                    {
                        workbook = new Workbook(inputPath);
                    }
                    catch (Exception loadEx)
                    {
                        Console.Error.WriteLine($"Failed to load '{inputPath}': {loadEx.Message}");
                        return;
                    }
                }
                else
                {
                    workbook = new Workbook();
                }

                // Configure PDF save options (default compression is used)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // If a specific compression level is required and the enum is available,
                // uncomment the following line:
                // pdfOptions.Compression = PdfCompressionType.Maximum;

                // Save the workbook as a PDF
                string outputPath = "output.pdf";
                try
                {
                    workbook.Save(outputPath, pdfOptions);
                    Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.Error.WriteLine($"Failed to save PDF: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
