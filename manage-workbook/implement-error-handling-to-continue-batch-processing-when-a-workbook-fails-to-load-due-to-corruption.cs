// Title: How to skip corrupted Excel workbooks and continue batch processing using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# skip corrupted .xlsx files during batch import | continue processing next workbook after load exception Aspose.Cells .NET | handle corrupted Excel workbook error while iterating files with Aspose.Cells | log and ignore invalid Excel files in Aspose.Cells batch operation C# | batch process multiple workbooks with Aspose.Cells and resilient error handling | skip unreadable Excel workbook and still save processed files using Aspose.Cells
// Tags: skip corrupted workbook Aspose.Cells | batch Excel processing error handling C# | Aspose.Cells load exception handling | continue processing after workbook load failure | log invalid Excel files Aspose.Cells | resilient workbook batch processing .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example shows how to iterate over a collection of Excel file paths, verify each file, and attempt to load it with Aspose.Cells. Load failures caused by corruption are caught, logged, and the loop continues to the next file. Successfully loaded workbooks are modified (e.g., adding a "Processed" sheet) and saved, with any save errors also logged, ensuring the batch operation proceeds despite individual workbook issues.
public class WorkbookBatchProcessor
{
    // Process a batch of Excel files. Corrupted files are skipped, processing continues.
    public void ProcessBatch(IEnumerable<string> inputFilePaths, string outputFolder)
    {
        // Ensure the output folder exists.
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        foreach (var inputPath in inputFilePaths)
        {
            // Verify the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                continue;
            }

            Workbook workbook = null;
            try
            {
                // Attempt to load the workbook.
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                // Log the error and continue with the next file.
                Console.WriteLine($"Failed to load workbook '{inputPath}': {ex.Message}");
                continue; // Skip to next file in the batch.
            }

            try
            {
                // Example operation: add a new worksheet named "Processed".
                Worksheet processedSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                processedSheet.Name = "Processed";
            }
            catch (Exception ex)
            {
                // Log modification errors but continue to attempt saving.
                Console.WriteLine($"Failed to modify workbook '{inputPath}': {ex.Message}");
            }

            // Save the processed workbook to the output folder.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Successfully processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any saving errors but continue processing other files.
                Console.WriteLine($"Failed to save workbook '{outputPath}': {ex.Message}");
            }
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Define input and output directories.
            string inputFolder = @"C:\InputExcels";
            string outputFolder = @"C:\ProcessedExcels";

            // Verify input folder exists.
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Get all .xlsx files from the input folder.
            var files = Directory.GetFiles(inputFolder, "*.xlsx");

            // Process the batch.
            var processor = new WorkbookBatchProcessor();
            processor.ProcessBatch(files, outputFolder);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors.
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
