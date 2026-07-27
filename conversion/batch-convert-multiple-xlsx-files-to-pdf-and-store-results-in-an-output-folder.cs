// Title: Batch Convert XLSX Files to PDF with Aspose.Cells ConversionUtility (C#)
// Description: A C# console utility that scans a given folder for *.xlsx workbooks, creates an output directory, and uses Aspose.Cells.Utility.ConversionUtility to convert each file to a PDF with the same name. Includes input validation, per‑file error handling, and simple command‑line usage.
// Keywords: Aspose.Cells batch conversion | C# XLSX to PDF | ConversionUtility Convert example | folder Excel to PDF | dotnet Excel PDF automation | command line Excel conversion | bulk Excel PDF generation | Aspose.Cells utility usage
// Common Searches: batch convert xlsx to pdf c# aspnet | aspocells convert multiple excel files to pdf | c# code to convert folder of xlsx files to pdf | aspocells ConversionUtility example | how to automate excel to pdf conversion in .net
// Developer Intent: Programmatically convert every .xlsx workbook in a specified input folder to a PDF in an output folder using Aspose.Cells.
// Use Cases: Nightly automation that archives daily Excel reports as PDFs for compliance. | Document‑management pipelines that ingest uploaded spreadsheets and store them as searchable PDFs. | A lightweight command‑line tool for end‑users to batch‑convert spreadsheets without opening Microsoft Excel.
// AI Prompts: Create a recursive version of the converter that preserves sub‑folder structure while converting all .xlsx files to PDF. | Write NUnit tests for ConvertFolder covering missing input folder, empty folder, and conversion failures. | Enhance the script to log each conversion (source, destination, status, timestamp) to a CSV file.

using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    // A C# console utility that scans a given folder for *.xlsx workbooks, creates an output directory, and uses Aspose.Cells.Utility.ConversionUtility to convert each file to a PDF with the same name. Includes input validation, per‑file error handling, and simple command‑line usage.
    public class BatchXlsxToPdfConverter
    {
        /// <param name="inputFolder">Folder containing source .xlsx files.</param>
        /// <param name="outputFolder">Folder where converted PDF files will be stored.</param>
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
                    // Build the destination PDF file path with the same base name
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Use Aspose.Cells.Utility.ConversionUtility to perform the conversion
                    ConversionUtility.Convert(sourcePath, destPath);

                    Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                }
            }
        }

        // Example entry point
        public static void Main(string[] args)
        {
            // Example usage:
            // args[0] = input folder, args[1] = output folder
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
