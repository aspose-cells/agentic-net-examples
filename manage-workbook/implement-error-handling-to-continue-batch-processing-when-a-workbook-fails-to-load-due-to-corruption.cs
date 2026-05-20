using System;
using System.Collections.Generic;
using Aspose.Cells;

class BatchProcessor
{
    // Processes a collection of Excel files.
    // If a file is corrupted, it logs the error and continues with the next file.
    public static void ProcessWorkbooks(IEnumerable<string> filePaths)
    {
        foreach (var path in filePaths)
        {
            try
            {
                // Load the workbook. Aspose.Cells will attempt to open the file.
                Workbook workbook = new Workbook(path);

                // Indicate that the workbook was opened in repair mode (safe mode).
                workbook.Settings.RepairLoad = true;

                // Example processing: calculate all formulas.
                workbook.CalculateFormula();

                // Save the processed workbook to an output folder.
                string outputFolder = "Processed";
                System.IO.Directory.CreateDirectory(outputFolder);
                string outputPath = System.IO.Path.Combine(outputFolder, System.IO.Path.GetFileName(path));
                workbook.Save(outputPath);

                Console.WriteLine($"Successfully processed and saved: {outputPath}");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
            {
                // Handle corrupted file without stopping the batch.
                Console.WriteLine($"Skipped corrupted file '{path}': {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors.
                Console.WriteLine($"Error processing '{path}': {ex.Message}");
            }
        }
    }

    static void Main()
    {
        // Example list of workbook file paths to process.
        var files = new List<string>
        {
            "Workbook1.xlsx",
            "CorruptedWorkbook.xlsx",
            "Workbook2.xlsx"
        };

        ProcessWorkbooks(files);
    }
}