using System;
using System.IO;
using Aspose.Cells;

namespace ExcelProcessor
{
    class ProcessExcelFiles
    {
        static void Main()
        {
            // Input and output directories
            string inputDirectory = @"C:\InputExcel";
            string outputDirectory = @"C:\OutputExcel";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Verify the input directory exists
            if (!Directory.Exists(inputDirectory))
            {
                Console.WriteLine($"Input directory not found: {inputDirectory}");
                return;
            }

            // Create a single LoadOptions instance for all workbooks
            LoadOptions loadOptions = new LoadOptions();

            // Get all .xlsx files in the input directory
            string[] files = Directory.GetFiles(inputDirectory, "*.xlsx");
            if (files.Length == 0)
            {
                Console.WriteLine("No .xlsx files found in the input directory.");
                return;
            }

            foreach (string sourcePath in files)
            {
                try
                {
                    // Ensure the source file exists before loading
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"File not found: {sourcePath}");
                        continue;
                    }

                    // Load the workbook using the shared LoadOptions
                    Workbook workbook = new Workbook(sourcePath, loadOptions);

                    // Determine the destination path
                    string destinationPath = Path.Combine(outputDirectory, Path.GetFileName(sourcePath));

                    // Save the workbook to the output directory
                    workbook.Save(destinationPath);
                    Console.WriteLine($"Processed: {sourcePath} -> {destinationPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{sourcePath}': {ex.Message}");
                }
            }
        }
    }
}