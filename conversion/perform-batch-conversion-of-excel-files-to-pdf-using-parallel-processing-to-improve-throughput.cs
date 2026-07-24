// Title: Parallel batch conversion of Excel workbooks to PDF with Aspose.Cells for .NET (C#)
// Description: Scans a folder for supported Excel formats, creates an output directory, and uses Aspose.Cells ConversionUtility inside Parallel.ForEach to convert each workbook to PDF concurrently. The sample logs each success, isolates failures, and runs without halting the whole batch.
// Keywords: Aspose.Cells batch conversion | Excel to PDF parallel C# | Convert multiple Excel files to PDF .NET | Aspose.Cells ConversionUtility example | bulk Excel PDF conversion | multi‑threaded Excel PDF generation | C# Aspose.Cells automation | high‑throughput Excel PDF conversion
// Common Searches: C# batch convert Excel files to PDF with Aspose.Cells | parallel processing for Excel to PDF conversion .NET | how to convert a folder of .xlsx to PDF using Aspose | Aspose.Cells bulk PDF export example | error handling in parallel Excel to PDF conversion
// Developer Intent: Convert every supported Excel workbook in a directory to PDF simultaneously using Aspose.Cells.
// Use Cases: Nightly automation that turns dozens of financial spreadsheets into PDF reports. | Large‑scale migration of legacy spreadsheets to PDF with maximum CPU utilization. | Web API that receives a zip of Excel files, converts each to PDF in parallel, and returns the results.
// AI Prompts: Generate C# code that adds progress reporting and cancellation tokens to the parallel Excel‑to‑PDF conversion using Aspose.Cells. | Show how to write conversion outcomes to a CSV log and retry failed files when using Parallel.ForEach with Aspose.Cells. | Explain how to control the degree of parallelism for batch Excel‑to‑PDF conversion in .NET to avoid exhausting system resources.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    // Scans a folder for supported Excel formats, creates an output directory, and uses Aspose.Cells ConversionUtility inside Parallel.ForEach to convert each workbook to PDF concurrently. The sample logs each success, isolates failures, and runs without halting the whole batch.
    public class BatchExcelToPdfConverter
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Directory containing source Excel files
            string sourceDirectory = "InputExcels";

            // Verify source directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory \"{sourceDirectory}\" does not exist.");
                return;
            }

            // Directory where PDF files will be saved
            string outputDirectory = "OutputPdfs";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Gather all files in the source directory
            string[] allFiles = Directory.GetFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly);

            // Define extensions that Aspose.Cells can convert
            var supportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".xlsx", ".xls", ".xlsm", ".xlsb", ".csv", ".ods", ".tsv"
            };

            // Filter only supported Excel files
            var filesToConvert = new List<string>();
            foreach (string filePath in allFiles)
            {
                if (supportedExtensions.Contains(Path.GetExtension(filePath)))
                {
                    filesToConvert.Add(filePath);
                }
            }

            // Perform conversion in parallel to improve throughput
            Parallel.ForEach(filesToConvert, sourcePath =>
            {
                try
                {
                    // Verify the source file still exists
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"Source file not found: {sourcePath}");
                        return;
                    }

                    // Build destination PDF file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputDirectory, fileNameWithoutExt + ".pdf");

                    // Convert Excel to PDF using Aspose.Cells utility
                    ConversionUtility.Convert(sourcePath, destPath);

                    Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
                }
                catch (Exception ex)
                {
                    // Log any conversion errors without stopping other tasks
                    Console.WriteLine($"Error converting {sourcePath}: {ex.Message}");
                }
            });
        }
    }
}
