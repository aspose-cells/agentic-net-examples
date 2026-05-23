using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    // Implements batch conversion of Excel files to another format
    // and reports percentage progress for each file.
    public class BatchConverter
    {
        // sourceDir : folder containing source Excel files
        // destDir   : folder where converted files will be saved
        // targetExt : desired output file extension (e.g., ".pdf", ".xlsx")
        public void Run(string sourceDir, string destDir, string targetExt)
        {
            try
            {
                // Validate input folders
                if (!Directory.Exists(sourceDir))
                {
                    Console.WriteLine($"Source directory not found: {sourceDir}");
                    return;
                }

                if (!Directory.Exists(destDir))
                    Directory.CreateDirectory(destDir);

                // Get all Excel files (any supported load format) in the source folder
                string[] sourceFiles = Directory.GetFiles(sourceDir, "*.*", SearchOption.TopDirectoryOnly);
                // Filter only files that Aspose.Cells can load (basic filter by extension)
                sourceFiles = Array.FindAll(sourceFiles, f =>
                    f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                    f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                    f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                    f.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) ||
                    f.EndsWith(".ods", StringComparison.OrdinalIgnoreCase));

                int totalFiles = sourceFiles.Length;
                if (totalFiles == 0)
                {
                    Console.WriteLine("No source Excel files found to convert.");
                    return;
                }

                for (int i = 0; i < totalFiles; i++)
                {
                    string srcPath = sourceFiles[i];
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(srcPath);
                    string destPath = Path.Combine(destDir, fileNameWithoutExt + targetExt);

                    // Ensure the source file exists before conversion
                    if (!File.Exists(srcPath))
                    {
                        Console.WriteLine($"Source file not found, skipping: {srcPath}");
                        continue;
                    }

                    try
                    {
                        // Perform conversion using Aspose.Cells utility method
                        ConversionUtility.Convert(srcPath, destPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting '{srcPath}': {ex.Message}");
                        continue;
                    }

                    // Calculate and display progress percentage
                    double percent = ((i + 1) * 100.0) / totalFiles;
                    Console.WriteLine($"Converted {i + 1}/{totalFiles} files ({percent:F2}%). Output: {destPath}");
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            try
            {
                // Define source and destination directories
                string sourceFolder = @"C:\InputExcelFiles";
                string destinationFolder = @"C:\ConvertedFiles";

                // Desired output format extension (e.g., ".pdf" for PDF)
                string targetExtension = ".pdf";

                // Run batch conversion
                var converter = new BatchConverter();
                converter.Run(sourceFolder, destinationFolder, targetExtension);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}