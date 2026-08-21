// Title: C# Parallel Batch Conversion of Excel (.xlsx) Files to PDF Using Aspose.Cells
// Description: Shows how to enumerate .xlsx files in a folder, ensure the target directory exists, and convert each workbook to PDF concurrently with Parallel.ForEach. The sample leverages Aspose.Cells Workbook.Save, handles missing files, logs successes and errors, and maximizes throughput on multi‑core .NET environments.
// Keywords: Aspose.Cells | C# | parallel processing | batch conversion | Excel to PDF | xlsx to pdf | high‑performance export | multi‑threaded PDF generation | dotnet example | GitHub sample | CLI utility
// Common Searches: parallel Excel to PDF conversion C# Aspose.Cells | batch convert .xlsx files to PDF with .NET | high throughput Excel PDF export Aspose | convert folder of Excel workbooks to PDF using Parallel.ForEach | Aspose.Cells example for bulk PDF generation
// Developer Intent: Convert a large set of Excel workbooks to PDF simultaneously to reduce overall processing time.
// Use Cases: Automated nightly job that archives all newly uploaded spreadsheets as PDFs. | Web service that receives multiple user spreadsheets and returns PDF versions in real time. | Command‑line tool for migrating a directory of legacy .xlsx reports to PDF on a multi‑core server.
// AI Prompts: Generate a C# method that accepts a list of Excel file paths and converts each to PDF with Aspose.Cells, including progress callbacks. | Explain how to control the degree of parallelism in the batch conversion to balance CPU usage and memory consumption. | Write unit tests that verify PDF output for valid files and proper error handling for missing or corrupted Excel files.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsBatchConversion
{
    // Shows how to enumerate .xlsx files in a folder, ensure the target directory exists, and convert each workbook to PDF concurrently with Parallel.ForEach. The sample leverages Aspose.Cells Workbook.Save, handles missing files, logs successes and errors, and maximizes throughput on multi‑core .NET environments.
    public class ExcelToPdfBatchConverter
    {
        // Converts a collection of Excel files to PDF in parallel.
        public static void ConvertFiles(IEnumerable<string> excelFilePaths, string outputFolder)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Process each file concurrently.
            Parallel.ForEach(excelFilePaths, excelPath =>
            {
                try
                {
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"Source file not found: {excelPath}");
                        return;
                    }

                    // Build the PDF file name based on the Excel file name.
                    string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Load the workbook and save as PDF using Aspose.Cells.
                    using (var workbook = new Workbook(excelPath))
                    {
                        workbook.Save(pdfPath, SaveFormat.Pdf);
                    }

                    Console.WriteLine($"Converted: {excelPath} -> {pdfPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting {excelPath}: {ex.Message}");
                }
            });
        }

        // Example entry point demonstrating usage.
        public static void Main()
        {
            try
            {
                // Directory containing Excel files to convert.
                string sourceDirectory = "InputExcels";

                // Verify source directory exists.
                if (!Directory.Exists(sourceDirectory))
                {
                    Console.WriteLine($"Source directory not found: {sourceDirectory}");
                    return;
                }

                // Retrieve all Excel files (you can adjust the pattern as needed).
                var excelFiles = Directory.GetFiles(sourceDirectory, "*.xlsx");

                if (excelFiles.Length == 0)
                {
                    Console.WriteLine($"No Excel files found in: {sourceDirectory}");
                    return;
                }

                // Directory where the resulting PDFs will be saved.
                string outputDirectory = "OutputPdfs";

                // Perform the batch conversion.
                ConvertFiles(excelFiles, outputDirectory);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
