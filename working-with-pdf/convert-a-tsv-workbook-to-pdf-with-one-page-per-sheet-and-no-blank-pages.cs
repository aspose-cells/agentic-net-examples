// Title: Convert a TSV workbook to PDF with one page per sheet and no blank pages using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a .tsv file into an Aspose.Cells Workbook, sets PdfSaveOptions.OnePagePerSheet = true, and saves it as a PDF. | Demonstrate how to add file‑existence checking and exception handling when converting a TSV file to PDF with Aspose.Cells. | Show how to configure Aspose.Cells PDF export so each worksheet is rendered on its own page and empty pages are avoided.
// Common Searches: asp.net aspose.cells convert tsv to pdf one page per worksheet | c# export tsv workbook to pdf without blank pages using aspose.cells | how to use PdfSaveOptions.OnePagePerSheet with tsv input in aspose.cells | load tsv file into workbook and save as pdf aspose.cells example
// Tags: Aspose.Cells PDF export OnePagePerSheet | TSV to PDF conversion Aspose.Cells C# | Workbook.Save with PdfSaveOptions .NET | File existence validation Aspose.Cells | Prevent blank pages PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the presence of an input TSV file, loads it into an Aspose.Cells Workbook using LoadFormat.Tsv, configures PdfSaveOptions.OnePagePerSheet to true to ensure each sheet appears on a separate PDF page, and saves the result as a PDF while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.tsv";
            const string outputPath = "output.pdf";

            // Verify that the input TSV file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the TSV file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options: one page per sheet.
            // Note: SkipEmptyRows/SkipEmptyColumns are not available in this version of Aspose.Cells.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF file successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
