// Title: C# – Batch load Excel workbooks with a shared LoadOptions using Aspose.Cells
// Description: Iterates over all files in a folder, filters supported Excel formats, loads each workbook with a single LoadOptions instance, optionally manipulates it, and saves the result to another directory while preserving the original format.
// Keywords: Aspose.Cells | C# batch load Excel | LoadOptions | process multiple workbooks | directory iteration | save workbook | Excel file conversion | Aspose.Cells .NET | bulk workbook processing
// Common Searches: Aspose.Cells load all Excel files in a folder | C# batch process workbooks with same LoadOptions | How to iterate through directory and save Excel files using Aspose.Cells | Bulk load Excel workbooks Aspose.Cells .NET | Apply identical LoadOptions to multiple workbooks
// Developer Intent: Load every Excel file in a specified folder with one LoadOptions object and write the workbooks to a target folder.
// Use Cases: Apply uniform calculation settings (e.g., disable formula evaluation) across many workbooks before analysis. | Convert a batch of CSV or legacy XLS files to modern XLSX while keeping the original layout. | Add or update a custom document property in all workbooks as part of a migration.
// AI Prompts: Write C# code that uses Aspose.Cells to batch load Excel files from a directory with LoadOptions that set LoadDataOnly = true, then save each workbook to a new folder. | Show how to add a custom document property to every workbook while iterating through a folder of Excel files with Aspose.Cells. | Provide a script that logs file name, worksheet count, and any errors to a CSV during batch processing with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    // Iterates over all files in a folder, filters supported Excel formats, loads each workbook with a single LoadOptions instance, optionally manipulates it, and saves the result to another directory while preserving the original format.
    public class BatchLoader
    {
        public static void Run()
        {
            // Directory containing the source Excel files
            string sourceDirectory = @"C:\InputExcelFiles";

            // Directory where processed files will be saved
            string outputDirectory = @"C:\ProcessedExcelFiles";

            // Verify source directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory does not exist: {sourceDirectory}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // Get all files in the source directory (filter later by extension)
            string[] files = Directory.GetFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly);

            // Common LoadOptions for all workbooks
            LoadOptions loadOptions = new LoadOptions();

            foreach (string filePath in files)
            {
                // Filter only supported Excel formats
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsb" && extension != ".csv")
                {
                    continue; // Skip unsupported files
                }

                // Verify the file actually exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook using the common LoadOptions
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Example operation: output basic info
                    Console.WriteLine($"Loaded '{Path.GetFileName(filePath)}' with {workbook.Worksheets.Count} worksheet(s).");

                    // Determine the output file path (same name, different folder)
                    string outputPath = Path.Combine(outputDirectory, Path.GetFileName(filePath));

                    // Save the workbook (preserving original format)
                    workbook.Save(outputPath);
                    Console.WriteLine($"Saved processed file to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                BatchLoader.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
