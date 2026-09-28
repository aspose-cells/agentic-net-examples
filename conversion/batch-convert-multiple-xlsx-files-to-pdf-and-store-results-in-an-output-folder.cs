// Title: How to batch convert XLSX files to PDF with Aspose.Cells ConversionUtility in C#
// AI Prompts: Write a C# console application that scans a given folder for *.xlsx files and uses Aspose.Cells ConversionUtility.Convert to create matching PDF files in a separate output directory. | Add try‑catch handling around each conversion call so that any failed file logs its exception but the loop continues processing the remaining workbooks. | Extend the program to traverse subfolders recursively and preserve the original folder hierarchy when saving the generated PDF files.
// Common Searches: c# batch convert all excel .xlsx files in a directory to pdf using aspose.cells | aspocells ConversionUtility convert multiple spreadsheets to pdf automatically | how to handle errors during bulk xlsx to pdf conversion in a C# console app
// Tags: batch xlsx to pdf conversion with Aspose.Cells | Aspose.Cells ConversionUtility for Excel to PDF | C# directory traversal for spreadsheet conversion | error handling in bulk Excel PDF export | recursive folder processing for Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    // // C# console program that iterates over .xlsx files in a specified input folder, converts each workbook to PDF using Aspose.Cells ConversionUtility, saves the PDFs to an output folder, and logs conversion errors without halting the batch process.
    public class BatchXlsxToPdfConverter
    {
        // Entry point
        public static void Main()
        {
            // Define the folder containing the source XLSX files
            string inputFolder = @"C:\InputXlsx";

            // Define the folder where the PDF files will be saved
            string outputFolder = @"C:\OutputPdf";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all .xlsx files in the input folder (non‑recursive)
            string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string sourcePath in xlsxFiles)
            {
                try
                {
                    // Build the destination PDF file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Convert the Excel file to PDF using Aspose.Cells ConversionUtility
                    ConversionUtility.Convert(sourcePath, destPath);

                    Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
                }
                catch (Exception ex)
                {
                    // Log any conversion errors but continue processing remaining files
                    Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
