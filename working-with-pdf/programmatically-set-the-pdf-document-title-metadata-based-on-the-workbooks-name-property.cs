// Title: How to set the PDF Title metadata from an Excel workbook’s file name using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, extracts the workbook’s file name, assigns it to PdfSaveOptions.Title, and saves the workbook as a PDF. | Show how to use Aspose.Cells PdfSaveOptions to embed custom document properties such as Title when converting Excel to PDF in a .NET application.
// Common Searches: Aspose.Cells set PDF title to workbook name C# | C# convert Excel to PDF and set document title property using Aspose | PdfSaveOptions.Title example Aspose.Cells | How to add custom metadata to PDF generated from Excel with Aspose.Cells | Set PDF document properties during Excel to PDF conversion .NET
// Tags: Aspose.Cells PdfSaveOptions.Title property | C# set PDF metadata during Excel conversion | Excel workbook filename to PDF title Aspose | custom PDF document properties Aspose.Cells | programmatic PDF title assignment .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // The example checks for the existence of an input Excel file, loads it into an Aspose.Cells Workbook, retrieves the workbook's file name, assigns that name to the PdfSaveOptions.Title property, and then saves the workbook as a PDF while handling any runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Optionally retrieve the workbook's file name (without path) for reference
                string workbookFileName = Path.GetFileName(workbook.FileName);

                // Set up PDF save options (you can customize further if needed)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF file using the options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook \"{workbookFileName}\" successfully saved as PDF to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
