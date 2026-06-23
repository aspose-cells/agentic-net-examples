using System;
using System.IO;
using Aspose.Cells;

namespace SharePointWorkbookEncryptionChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Folder that contains the files to be checked.
                // Replace with your actual path or obtain it from args.
                string inputFolder = @"C:\InputFiles";

                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder does not exist: {inputFolder}");
                    return;
                }

                // Enumerate all files in the folder (non‑recursive; modify if needed).
                string[] filePaths = Directory.GetFiles(inputFolder);
                if (filePaths.Length == 0)
                {
                    Console.WriteLine("No files found in the input folder.");
                    return;
                }

                foreach (string filePath in filePaths)
                {
                    try
                    {
                        // Ensure the file exists before processing.
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found: {filePath}");
                            continue;
                        }

                        // Open the file as a stream.
                        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                        {
                            // Detect file format and encryption status using Aspose.Cells.
                            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(fs);
                            bool isEncrypted = formatInfo.IsEncrypted;

                            // Log the result.
                            Console.WriteLine($"{Path.GetFileName(filePath)}: Encrypted = {isEncrypted}");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Handle errors for individual files without stopping the whole process.
                        Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Global exception handler.
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}