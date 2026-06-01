using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    public class XlsxToPdfBatchConverter
    {
        /// <summary>
        /// Converts all .xlsx files in the specified input folder to PDF files in the output folder.
        /// </summary>
        /// <param name="inputFolder">Folder containing source .xlsx files.</param>
        /// <param name="outputFolder">Folder where converted PDF files will be saved.</param>
        public static void ConvertFolder(string inputFolder, string outputFolder)
        {
            // Ensure the input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Create the output directory if it does not exist
            Directory.CreateDirectory(outputFolder);

            // Get all .xlsx files (case‑insensitive) in the input folder
            string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string sourcePath in xlsxFiles)
            {
                try
                {
                    // Build the destination PDF file path (same file name, .pdf extension)
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Perform the conversion using Aspose.Cells.Utility.ConversionUtility
                    ConversionUtility.Convert(sourcePath, destPath);

                    Console.WriteLine($"Converted: {Path.GetFileName(sourcePath)} -> {Path.GetFileName(destPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }

        // Example entry point
        public static void Main(string[] args)
        {
            // Example usage:
            // args[0] = input folder, args[1] = output folder
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: XlsxToPdfBatchConverter <inputFolder> <outputFolder>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];

            ConvertFolder(inputFolder, outputFolder);
        }
    }
}