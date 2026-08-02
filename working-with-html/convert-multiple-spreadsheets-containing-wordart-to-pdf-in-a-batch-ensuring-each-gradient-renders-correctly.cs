// Title: Batch convert Excel files with WordArt to PDF (preserve gradients) using Aspose.Cells for .NET
// Description: C# utility that scans a folder for Excel workbooks, loads each file with Aspose.Cells, and saves it as a PDF while keeping WordArt objects and their gradient fills intact. Includes folder validation, command‑line support, robust error handling, and console logging for easy automation.
// Keywords: Aspose.Cells batch PDF conversion | Excel WordArt to PDF .NET | preserve gradient WordArt PDF | C# convert multiple Excel files | command line Excel to PDF tool | Aspose.Cells PDF save options | automated Excel PDF export
// Common Searches: batch convert Excel to PDF with WordArt Aspose.Cells | C# preserve gradient WordArt when exporting Excel to PDF | convert folder of .xlsx files to PDF using Aspose.Cells | command line tool for Excel to PDF conversion .NET | Aspose.Cells example for batch PDF export
// Developer Intent: Automatically convert a directory of Excel workbooks that contain WordArt graphics into PDF files, ensuring the original gradient styling of the WordArt is retained.
// Use Cases: Nightly generation of PDF reports from Excel templates that include WordArt branding. | Command‑line utility for end‑users to batch‑convert spreadsheets to PDF without manual intervention. | Integration into a larger document‑processing pipeline that creates PDF invoices from Excel sheets with decorative WordArt.
// AI Prompts: Write C# code using Aspose.Cells to batch convert Excel files with WordArt to PDF, preserving gradient fills and handling errors gracefully. | Show how to set Aspose.Cells PDF save options to ensure WordArt gradients render correctly in the output PDF. | Create a PowerShell script that runs the compiled .NET batch converter, accepting input and output folder paths as parameters.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchConversion
{
    // C# utility that scans a folder for Excel workbooks, loads each file with Aspose.Cells, and saves it as a PDF while keeping WordArt objects and their gradient fills intact. Includes folder validation, command‑line support, robust error handling, and console logging for easy automation.
    public class WordArtToPdfBatchConverter
    {
        /// <param name="inputFolder">Folder containing source Excel files (e.g., .xlsx, .xls).</param>
        /// <param name="outputFolder">Folder where the resulting PDF files will be saved.</param>
        public static void ConvertFolder(string inputFolder, string outputFolder)
        {
            // Validate input folder
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Supported Excel extensions
            string[] extensions = new[] { "*.xlsx", "*.xls", "*.xlsm", "*.xlsb", "*.csv" };

            foreach (string pattern in extensions)
            {
                foreach (string sourcePath in Directory.GetFiles(inputFolder, pattern, SearchOption.TopDirectoryOnly))
                {
                    // Verify source file exists (defensive check)
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"Source file not found, skipping: {sourcePath}");
                        continue;
                    }

                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    try
                    {
                        // Load workbook and save as PDF
                        Workbook wb = new Workbook(sourcePath);
                        wb.Save(destPath, SaveFormat.Pdf);
                        Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                    }
                }
            }
        }

        // Example usage
        public static void Run()
        {
            string inputDir = @"C:\InputExcelFiles";
            string outputDir = @"C:\OutputPdfFiles";

            try
            {
                ConvertFolder(inputDir, outputDir);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            // Optionally allow command‑line arguments for folders
            if (args.Length >= 2)
            {
                WordArtToPdfBatchConverter.ConvertFolder(args[0], args[1]);
            }
            else
            {
                WordArtToPdfBatchConverter.Run();
            }
        }
    }
}
