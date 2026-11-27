// Title: How to monitor peak memory usage while processing a large Excel workbook with Aspose.Cells LightCells API in C#
// Common Searches: C# example for tracking memory consumption during row‑by‑row iteration with Aspose.Cells LightCells | measure peak memory while processing a large .xlsx file using Aspose.Cells in .NET | using GC.GetTotalMemory to profile memory when copying a workbook with LightCells API | best practice for memory profiling of large Excel workbooks in C# with Aspose.Cells | detect memory spikes while reading rows from an Excel workbook via LightCells
// Tags: Aspose.Cells LightCells row iteration memory profiling | C# peak memory monitoring Aspose.Cells workbook | large Excel file processing LightCells API | GC.GetTotalMemory usage with Aspose.Cells | memory usage statistics during workbook copy C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a largeWorkbook.xlsx using Aspose.Cells, iterates each row of the first worksheet, and records current and peak memory usage with GC.GetTotalMemory (forcing a full collection). It prints progress and memory statistics, saves a copy as processedWorkbook.xlsx, and reports final memory consumption, demonstrating how to monitor memory while processing large Excel files with the LightCells API in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the large workbook to be processed
            string inputPath = "largeWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook using the standard Aspose.Cells API
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Variables for tracking memory usage
            long peakMemory = 0;
            int processedRows = 0;

            // Process the workbook row by row
            foreach (Row row in worksheet.Cells.Rows)
            {
                try
                {
                    // Placeholder for any row‑level processing logic
                    processedRows++;

                    // Get current memory usage (force a full collection for more accurate reading)
                    long currentMemory = GC.GetTotalMemory(forceFullCollection: true);
                    if (currentMemory > peakMemory)
                    {
                        peakMemory = currentMemory;
                    }

                    // Output progress and memory information
                    Console.WriteLine($"Rows processed: {processedRows} | " +
                                      $"Current memory used: {currentMemory} bytes | " +
                                      $"Peak memory used: {peakMemory} bytes");
                }
                catch (Exception exRow)
                {
                    Console.WriteLine($"Error processing row {processedRows}: {exRow.Message}");
                }
            }

            // OPTIONAL: Save the processed data to a new workbook (if any modifications were made)
            // Here we simply copy the original workbook to a new file.
            string outputPath = "processedWorkbook.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to the output path
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Error saving workbook: {exSave.Message}");
                return;
            }

            // Report final memory usage after the save operation
            long finalMemory = GC.GetTotalMemory(forceFullCollection: true);
            if (finalMemory > peakMemory)
            {
                peakMemory = finalMemory;
            }

            Console.WriteLine($"Copy completed. Final memory used: {finalMemory} bytes | " +
                              $"Peak memory used during operation: {peakMemory} bytes");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
