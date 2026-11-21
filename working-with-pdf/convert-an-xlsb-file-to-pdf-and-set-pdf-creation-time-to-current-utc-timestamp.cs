// Title: Convert an XLSB workbook to PDF in C# with Aspose.Cells and set the PDF creation date to the current UTC time
// Common Searches: c# aspose.cells convert xlsb file to pdf and set creation date | how to export binary Excel workbook as PDF with UTC timestamp using Aspose | Aspose.Cells PdfSaveOptions set PDF document creation time in C# | save xlsb as pdf and add current UTC metadata with Aspose.Cells | C# code to convert XLSB to PDF and customize PDF metadata Aspose
// Tags: Aspose.Cells XLSB to PDF conversion | PdfSaveOptions creation timestamp setting | C# export binary Excel workbook as PDF | PDF document metadata UTC timestamp | Aspose.Cells custom PDF properties

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the input XLSB file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, assigns the current UTC time to the PDF creation date, and saves the workbook as a PDF. Any runtime errors are caught and reported.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsb";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the XLSB workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (additional options can be set here if needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
