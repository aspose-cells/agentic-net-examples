// Title: Handle corrupted Excel workbooks in a batch process with Aspose.Cells for C#
// Common Searches: C# Aspose.Cells skip corrupted workbook and continue batch processing | How to catch load errors for multiple Excel files using Aspose.Cells .NET | Continue processing remaining workbooks after a file fails to open with Aspose.Cells | Batch load Excel files with Aspose.Cells and ignore corrupted files in C# | Aspose.Cells exception handling for corrupted .xlsx during bulk conversion
// Tags: Aspose.Cells workbook load exception handling | skip corrupted .xlsx files C# | batch processing Excel workbooks Aspose.Cells | continue processing after workbook load failure | error logging for Aspose.Cells batch operation

using System;
using System.IO;
using Aspose.Cells;

// The example creates the output directory if needed, iterates through a list of Excel file paths, checks each file's existence, attempts to load each workbook with Aspose.Cells, saves a copy to the output folder, and catches any exceptions (such as file corruption) to log the error and continue processing the remaining workbooks.
public class WorkbookBatchProcessor
{
    /// <param name="inputFiles">Array of full paths to the workbook files to process.</param>
    /// <param name="outputFolder">Folder where processed workbooks will be saved.</param>
    public void ProcessWorkbooks(string[] inputFiles, string outputFolder)
    {
        // Ensure the output directory exists.
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        foreach (string filePath in inputFiles)
        {
            // Verify the file exists before attempting to load.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: '{filePath}'. Skipping.");
                continue;
            }

            try
            {
                // Load workbook.
                Workbook workbook = new Workbook(filePath);

                // Example processing: simply save a copy to the output folder.
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Log the error and continue with the next workbook.
                Console.WriteLine($"Failed to process '{filePath}': {ex.Message}");
            }
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Expect at least two arguments: output folder and one or more input files.
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <outputFolder> <inputFile1> [<inputFile2> ...]");
            return;
        }

        string outputFolder = args[0];
        string[] inputFiles = new string[args.Length - 1];
        Array.Copy(args, 1, inputFiles, 0, inputFiles.Length);

        try
        {
            WorkbookBatchProcessor processor = new WorkbookBatchProcessor();
            processor.ProcessWorkbooks(inputFiles, outputFolder);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
