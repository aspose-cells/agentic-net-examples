// Title: C# Batch Convert XLSX Files to PDF Using Aspose.Cells ConversionUtility
// Description: A command‑line C# utility that scans a specified input folder, validates or creates the output directory, and uses Aspose.Cells.Utility.ConversionUtility.Convert to transform each .xlsx workbook into a PDF with the same base name. The program logs progress, handles missing folders and conversion errors, and is ready for integration into scripts or CI pipelines.
// Keywords: Aspose.Cells batch conversion | C# XLSX to PDF | ConversionUtility example | convert folder of Excel files | command line Excel to PDF | Aspose.Cells .NET PDF export | GitHub Aspose.Cells sample
// Common Searches: batch convert xlsx to pdf c# aspose.cells | asp.net convert all excel files in a folder to pdf | aspose.cells conversionutility convert multiple workbooks | command line tool to convert excel to pdf using asp.net | c# code example for folder based excel pdf conversion
// Developer Intent: Automatically convert every .xlsx file in a given directory to a PDF and save the results in a designated output folder.
// Use Cases: Automate nightly report archiving by turning generated Excel sheets into PDFs. | Create a lightweight CLI tool for end‑users to batch‑convert spreadsheet templates before distribution. | Integrate Excel‑to‑PDF conversion into CI/CD pipelines to validate visual output of workbook reports.
// AI Prompts: Write C# code that uses Aspose.Cells ConversionUtility to batch convert all Excel files in a folder to PDFs, including folder checks and exception handling. | Generate a PowerShell script that runs the compiled BatchXlsxToPdfConverter with input and output path arguments. | Explain how to modify the example to recursively process subfolders while preserving the original directory hierarchy in the PDF output.

using System;
using System.IO;
using Aspose.Cells.Utility;   // Required for ConversionUtility

namespace AsposeCellsBatchConversion
{
    // A command‑line C# utility that scans a specified input folder, validates or creates the output directory, and uses Aspose.Cells.Utility.ConversionUtility.Convert to transform each .xlsx workbook into a PDF with the same base name. The program logs progress, handles missing folders and conversion errors, and is ready for integration into scripts or CI pipelines.
    public class BatchXlsxToPdfConverter
    {
        /// <param name="inputFolder">Folder containing source .xlsx files.</param>
        /// <param name="outputFolder">Folder where converted PDF files will be stored.</param>
        public static void ConvertFolder(string inputFolder, string outputFolder)
        {
            // Ensure the input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Create the output folder if it does not exist
            Directory.CreateDirectory(outputFolder);

            // Get all .xlsx files in the input folder (non‑recursive)
            string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            if (xlsxFiles.Length == 0)
            {
                Console.WriteLine("No .xlsx files found to convert.");
                return;
            }

            foreach (string sourcePath in xlsxFiles)
            {
                try
                {
                    // Build the destination PDF file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Use Aspose.Cells.Utility.ConversionUtility to perform the conversion
                    // This follows the provided rule: ConversionUtility.Convert(string, string)
                    ConversionUtility.Convert(sourcePath, destPath);

                    Console.WriteLine($"Converted: {Path.GetFileName(sourcePath)} -> {Path.GetFileName(destPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting file '{sourcePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }

        // Example entry point
        public static void Main(string[] args)
        {
            // Example usage:
            // args[0] = input folder path, args[1] = output folder path
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: BatchXlsxToPdfConverter <inputFolder> <outputFolder>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];

            ConvertFolder(inputFolder, outputFolder);
        }
    }
}
