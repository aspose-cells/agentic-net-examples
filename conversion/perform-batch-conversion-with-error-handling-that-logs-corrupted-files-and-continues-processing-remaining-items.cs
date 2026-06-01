using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchConversionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example input files (replace with actual paths)
            string[] sourceFiles = new string[]
            {
                "File1.xlsx",
                "File2.xls",
                "CorruptedFile.xlsx", // This file is expected to be corrupted for demo
                "File3.csv"
            };

            // Destination directory for converted PDFs
            string outputDirectory = "ConvertedFiles";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Perform batch conversion
            BatchConvertToPdf(sourceFiles, outputDirectory);
        }

        /// <summary>
        /// Converts a list of Excel-related files to PDF.
        /// Corrupted files are logged and the process continues with remaining files.
        /// </summary>
        /// <param name="sourceFiles">Array of source file paths.</param>
        /// <param name="outputDir">Directory where converted PDFs will be saved.</param>
        static void BatchConvertToPdf(string[] sourceFiles, string outputDir)
        {
            foreach (string sourcePath in sourceFiles)
            {
                try
                {
                    // Verify source file exists
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"Source file not found: {sourcePath}");
                        continue;
                    }

                    // Determine output file name (same name with .pdf extension)
                    string outputFileName = Path.GetFileNameWithoutExtension(sourcePath) + ".pdf";
                    string destPath = Path.Combine(outputDir, outputFileName);

                    // Load options (default, can be customized if needed)
                    LoadOptions loadOptions = new LoadOptions();

                    // Create PDF save options and enable IgnoreError to hide rendering errors
                    PdfSaveOptions saveOptions = new PdfSaveOptions
                    {
                        // PaginatedSaveOptions.IgnoreError property
                        IgnoreError = true
                    };

                    // Perform conversion using the utility method that accepts options
                    ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

                    Console.WriteLine($"Successfully converted: {sourcePath} -> {destPath}");
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
                {
                    // Specific handling for corrupted files
                    Console.WriteLine($"Corrupted file detected and skipped: {sourcePath}");
                    // Optionally log details to a file or monitoring system here
                }
                catch (Exception ex)
                {
                    // General error handling for other unexpected issues
                    Console.WriteLine($"Error converting file '{sourcePath}': {ex.Message}");
                    // Continue with next file
                }
            }
        }
    }
}