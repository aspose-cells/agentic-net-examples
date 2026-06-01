using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    public class ExcelToPdfBatchConverter
    {
        // Entry point for the batch conversion
        public static void Run(string sourceFolder, string destinationFolder)
        {
            try
            {
                // Validate input folder
                if (string.IsNullOrWhiteSpace(sourceFolder) || !Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                    return;
                }

                // Ensure output folder exists
                Directory.CreateDirectory(destinationFolder ?? string.Empty);

                // Gather all Excel files (common extensions) from the source folder and subfolders
                string[] excelExtensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb", ".csv", ".ods", ".xlsxml" };
                List<string> filesToConvert = new List<string>();

                foreach (string ext in excelExtensions)
                {
                    filesToConvert.AddRange(Directory.GetFiles(sourceFolder, $"*{ext}", SearchOption.AllDirectories));
                }

                if (filesToConvert.Count == 0)
                {
                    Console.WriteLine("No Excel files found to convert.");
                    return;
                }

                Console.WriteLine($"Found {filesToConvert.Count} file(s). Starting parallel conversion...");

                // Perform conversion in parallel
                ParallelOptions parallelOptions = new ParallelOptions
                {
                    // Adjust degree of parallelism as needed; default uses as many threads as processors
                    MaxDegreeOfParallelism = Environment.ProcessorCount
                };

                Parallel.ForEach(filesToConvert, parallelOptions, sourcePath =>
                {
                    try
                    {
                        // Verify the source file still exists
                        if (!File.Exists(sourcePath))
                        {
                            Console.WriteLine($"Source file not found (skipped): {sourcePath}");
                            return;
                        }

                        // Determine relative path to preserve folder structure in output
                        string relativePath = Path.GetRelativePath(sourceFolder, sourcePath);
                        string outputPdfPath = Path.Combine(destinationFolder,
                            Path.ChangeExtension(relativePath, ".pdf"));

                        // Ensure the output directory exists
                        string? outputDir = Path.GetDirectoryName(outputPdfPath);
                        Directory.CreateDirectory(outputDir ?? string.Empty);

                        // Convert Excel to PDF using Aspose.Cells utility
                        ConversionUtility.Convert(sourcePath, outputPdfPath);

                        Console.WriteLine($"Converted: {sourcePath} -> {outputPdfPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                    }
                });

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }

        // Simple console entry point to invoke the batch converter
        public static void Main(string[] args)
        {
            // Expect two arguments: source folder and destination folder
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsBatchConversion <sourceFolder> <destinationFolder>");
                return;
            }

            string sourceFolder = args[0];
            string destinationFolder = args[1];

            Run(sourceFolder, destinationFolder);
        }
    }
}