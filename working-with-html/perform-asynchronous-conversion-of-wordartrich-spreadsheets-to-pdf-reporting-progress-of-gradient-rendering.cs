// Title: Asynchronously convert an Excel workbook with WordArt to PDF and report progress using Aspose.Cells for .NET
// AI Prompts: Write a C# async method that loads an .xlsx file containing WordArt, saves it as PDF on a background thread, and reports the conversion progress through IProgress<int>. | Implement robust error handling for an asynchronous Excel‑to‑PDF conversion that validates the input path, creates missing output directories, and wraps any thrown exceptions in an AggregateException. | Show how to use Aspose.Cells PdfSaveOptions with a background thread to keep the UI responsive while converting a workbook that includes graphic objects to PDF.
// Common Searches: c# aspnet convert excel file with wordart to pdf asynchronously using aspose.cells | how to use IProgress<int> to track Aspose.Cells workbook.Save progress | run Aspose.Cells workbook.Save on Task.Run to keep UI responsive | aspose.cells pdfsaveoptions async conversion example with progress reporting | handle file not found and create output folder when converting excel to pdf in .net
// Tags: background excel-to-pdf conversion aspnet | wordart rendering pdfsaveoptions aspose.cells | progress reporting iprogress workbook.save | task.run non‑blocking conversion aspose.cells | aggregateexception exception wrapping

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample validates the source Excel file, creates the destination folder if needed, loads the workbook, configures PdfSaveOptions, reports start (0%) and finish (100%) via IProgress<int>, executes workbook.Save on a background thread to avoid blocking, and wraps any errors in an AggregateException for proper async error handling.
public class SpreadsheetPdfConverter
{
    /// <param name="inputFile">Full path to the source Excel file.</param>
    /// <param name="outputFile">Full path where the PDF will be saved.</param>
    /// <param name="progress">Progress reporter (percentage 0‑100) for the save operation.</param>
    /// <returns>A task that completes when the conversion finishes.</returns>
    public async Task ConvertToPdfAsync(string inputFile, string outputFile, IProgress<int> progress)
    {
        if (!File.Exists(inputFile))
            throw new FileNotFoundException($"Input file not found: {inputFile}");

        try
        {
            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the workbook.
            var workbook = new Workbook(inputFile);

            // Configure PDF save options (no progress handler in this version).
            var pdfOptions = new PdfSaveOptions();

            // Report start of the operation.
            progress?.Report(0);

            // Perform the save operation on a background thread to avoid blocking.
            await Task.Run(() =>
            {
                workbook.Save(outputFile, pdfOptions);
            });

            // Report completion.
            progress?.Report(100);
        }
        catch (Exception ex)
        {
            // Wrap the exception for async context.
            throw new AggregateException("Error converting Excel to PDF.", ex);
        }
    }
}

public class Program
{
    public static async Task Main(string[] args)
    {
        var converter = new SpreadsheetPdfConverter();
        var progress = new Progress<int>(p => Console.WriteLine($"Save progress: {p}%"));

        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            await converter.ConvertToPdfAsync(inputPath, outputPath, progress);
            Console.WriteLine("Conversion completed successfully.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine(fnfEx.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Conversion failed: {ex.Message}");
        }
    }
}
