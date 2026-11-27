// Title: Create a PDF from an XLSX workbook with visible gridlines using Aspose.Cells for .NET
// Common Searches: Aspose.Cells how to include gridlines when saving workbook as PDF in C# | C# convert Excel to PDF with gridlines displayed using Aspose.Cells | PdfSaveOptions GridlineType setting for visible gridlines in Aspose.Cells .NET | Render dotted gridlines in PDF output from XLSX with Aspose.Cells | OnePagePerSheet option with gridlines in Aspose.Cells PDF conversion
// Tags: Aspose.Cells PDFSaveOptions GridlineType | C# Excel to PDF conversion with gridlines | visible gridlines PDF rendering Aspose.Cells | OnePagePerSheet PDF option Aspose.Cells | handling missing input file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace SolidGridlinesPdfDemo
{
    // The example loads an XLSX workbook, configures PdfSaveOptions to render visible (dotted) gridlines and to place each worksheet on a separate PDF page, then saves the workbook as a PDF while checking for missing input files and handling exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX workbook
            const string inputPath = "input.xlsx";
            // Path for the generated PDF
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.Error.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                    return;
                }

                // Load the workbook from the XLSX file
                Workbook workbook = new Workbook(inputPath);

                // Create PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Render gridlines (the enum does not have a 'Solid' value; use Dotted for visible lines)
                    GridlineType = GridlineType.Dotted,
                    // Optional: put each worksheet on a separate PDF page
                    OnePagePerSheet = true
                };

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF successfully created with gridlines: \"{outputPath}\"");
            }
            catch (Exception ex)
            {
                // Output any unexpected errors
                Console.Error.WriteLine("An error occurred while generating the PDF:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
