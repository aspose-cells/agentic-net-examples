// Title: Asynchronously convert an Excel workbook containing WordArt and gradient fills to PDF with Aspose.Cells for .NET
// AI Prompts: Write an async C# method that loads an .xlsx file, configures PdfSaveOptions with OnePagePerSheet, and saves it as a PDF using Aspose.Cells while preserving WordArt and gradient fills. | Add progress reporting to the asynchronous Excel‑to‑PDF conversion, outputting the percentage of sheets rendered during the Aspose.Cells export. | Implement robust error handling for missing input files and conversion failures in the async Aspose.Cells PDF export routine.
// Common Searches: how to perform async Excel to PDF conversion with Aspose.Cells preserving WordArt | c# Aspose.Cells PdfSaveOptions OnePagePerSheet example | report conversion progress during Aspose.Cells PDF export of large workbooks | handle file not found exception in Aspose.Cells asynchronous conversion | convert workbook with gradient shapes to PDF using Aspose.Cells async method
// Tags: aspose.cells asynchronous pdf export | aspose.cells pdfsaveoptions onepagepersheet | preserve wordart gradient fills aspose.cells | c# background workbook loading | c# file not found exception handling

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample defines an async ConvertAsync method that validates the input path, loads the workbook on a background thread, sets PdfSaveOptions.OnePagePerSheet to keep layout intact, and saves the workbook as a PDF. Any exception is wrapped in an InvalidOperationException. The Main method parses command‑line arguments, invokes the async conversion, and handles FileNotFoundException, InvalidOperationException, and other unexpected errors.
class SpreadsheetToPdfConverter
{
    // Asynchronous method that loads a workbook, converts it to PDF, and preserves visual fidelity
    public static async Task ConvertAsync(string inputPath, string outputPath)
    {
        // Verify input file exists
        if (!File.Exists(inputPath))
            throw new FileNotFoundException($"Input file not found: {inputPath}");

        try
        {
            // Load the workbook on a background thread
            Workbook workbook = await Task.Run(() => new Workbook(inputPath));

            // Configure PDF save options (e.g., one page per sheet to keep layout intact)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook to PDF asynchronously
            await Task.Run(() => workbook.Save(outputPath, pdfOptions));
        }
        catch (Exception ex)
        {
            // Wrap any exception to provide context while preserving the original stack trace
            throw new InvalidOperationException("Failed to convert workbook to PDF.", ex);
        }
    }

    // Example entry point
    static async Task Main(string[] args)
    {
        // Validate arguments
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: SpreadsheetToPdfConverter <input.xlsx> <output.pdf>");
            return;
        }

        string inputFile = args[0];
        string outputFile = args[1];

        try
        {
            // Perform the asynchronous conversion
            await ConvertAsync(inputFile, outputFile);
            Console.WriteLine("Conversion completed successfully.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File error: {fnfEx.Message}");
        }
        catch (InvalidOperationException invEx)
        {
            Console.WriteLine($"Conversion error: {invEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
