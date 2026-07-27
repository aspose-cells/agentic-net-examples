using System;
using System.IO;
using Aspose.Cells;

public class BatchWorkbookProcessor
{
    // Processes a collection of Excel files, attempting to load each one.
    // If a file is corrupted, the error is logged and processing continues with the next file.
    public void ProcessWorkbooks(string[] inputFiles, string outputDirectory)
    {
        // Ensure the output directory exists
        Directory.CreateDirectory(outputDirectory);

        foreach (string filePath in inputFiles)
        {
            try
            {
                // Attempt to load the workbook
                Workbook workbook = new Workbook(filePath);

                // Optional: indicate that the workbook was loaded in repair mode (if needed later)
                workbook.Settings.RepairLoad = true;

                // Example processing: simply save a copy to the output folder
                string outputPath = Path.Combine(outputDirectory, Path.GetFileName(filePath));
                workbook.Save(outputPath);

                Console.WriteLine($"Successfully processed: {filePath}");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
            {
                // Specific handling for corrupted files – log and continue
                Console.WriteLine($"Corrupted file skipped: {filePath} (Reason: {ex.Message})");
            }
            catch (Exception ex)
            {
                // General error handling – log and continue
                Console.WriteLine($"Error processing file {filePath}: {ex.Message}");
            }
        }
    }
}

// Example usage
public class Program
{
    public static void Main()
    {
        string[] filesToProcess = new string[]
        {
            @"C:\Data\Workbook1.xlsx",
            @"C:\Data\Workbook2.xlsx",
            @"C:\Data\CorruptedWorkbook.xlsx",
            @"C:\Data\Workbook3.xlsx"
        };

        string outputFolder = @"C:\Data\Processed";

        BatchWorkbookProcessor processor = new BatchWorkbookProcessor();
        processor.ProcessWorkbooks(filesToProcess, outputFolder);
    }
}