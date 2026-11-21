// Title: Load a TSV file into Aspose.Cells, enable OnePagePerSheet, and save as PDF with default security in C#
// Common Searches: C# Aspose.Cells load TSV workbook and export to PDF with one page per sheet | How to use PdfSaveOptions OnePagePerSheet when converting TSV to PDF in .NET | Aspose.Cells C# convert tab-separated values to PDF preserving sheet layout | Saving a workbook loaded from TSV as PDF with default security using Aspose.Cells | LoadOptions for TSV format in Aspose.Cells C# example
// Tags: Aspose.Cells TSV LoadOptions | PdfSaveOptions OnePagePerSheet C# | Export workbook to PDF default security Aspose.Cells | C# convert TSV to PDF Aspose.Cells | Workbook.Save PDF page fitting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input TSV file, loads it into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Tsv, configures PdfSaveOptions to fit each sheet on a single page, ensures the output folder exists, and saves the workbook as a PDF with default security.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source TSV file
            string inputPath = "input.tsv";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Path for the resulting PDF file
            string outputPath = "output.pdf";

            // Load the TSV workbook using LoadOptions to specify the format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options to fit each sheet onto a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
