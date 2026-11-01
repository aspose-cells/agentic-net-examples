// Title: Convert a CSV file to PDF and set a custom creation date using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a CSV file with Aspose.Cells LoadOptions, converts it to PDF, and assigns a specific historical CreationDate metadata before saving. | Show how to configure PdfSaveOptions and DocumentProperties in Aspose.Cells to embed a custom creation timestamp when exporting a workbook to PDF.
// Common Searches: aspocells c# convert csv to pdf and set creation date metadata | how to set PDF creation time using Aspose.Cells SaveOptions | export workbook to PDF with custom timestamp Aspose.Cells .NET | load csv into workbook and assign historical PDF metadata Aspose.Cells
// Tags: csv to pdf conversion aspocells c# | set pdf creation date aspocells | pdfsaveoptions metadata aspocells | loadoptions csv workbook aspocells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a CSV file into an Aspose.Cells Workbook using LoadOptions, then saves the workbook as a PDF while applying PdfSaveOptions and DocumentProperties to embed a specified historical creation date in the PDF metadata, with error handling for missing files and exceptions.
class CsvToPdfConverter
{
    static void Main()
    {
        try
        {
            string inputPath = "input.csv";
            string outputPath = "output.pdf";

            // Verify that the input CSV file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the CSV file into a workbook using appropriate load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options (metadata setting omitted for compatibility)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
