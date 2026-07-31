// Title: C# CLI to batch convert chart‑free Excel workbooks to PDF with Aspose.Cells
// Description: A console application that receives a folder path, scans for Excel files, skips any workbook containing charts, and converts the remaining files to PDF using Aspose.Cells' ConversionUtility. The PDF is saved beside the source file with the same base name.
// Keywords: Aspose.Cells | Excel to PDF conversion | C# console app | batch Excel PDF | skip chart worksheets | folder processing | command line utility | .NET PDF generation | ConversionUtility | automated report export
// Common Searches: C# command line batch convert Excel to PDF Aspose.Cells | how to ignore Excel files with charts when converting to PDF | convert all .xlsx files in a directory to PDF using .NET | Aspose.Cells example for folder‑level Excel PDF conversion | skip chart‑containing worksheets during PDF export
// Developer Intent: Create a .NET console tool that iterates through a directory, filters Excel workbooks, excludes those with charts, and generates PDF files for the rest.
// Use Cases: Nightly automation that turns chart‑free Excel templates into PDF reports for distribution. | Archival script that converts incoming data spreadsheets to PDF while leaving chart‑rich files untouched. | CI/CD step that validates Excel inputs can be rendered as PDF without chart rendering errors.
// AI Prompts: Add structured logging and a progress bar to the Excel‑to‑PDF CLI while keeping the chart‑skip logic intact. | Write an xUnit test that verifies files containing charts are not converted to PDF by the folder processor. | Modify the utility to output a JSON summary with counts of converted, skipped, and failed files.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace ExcelToPdfConverter
{
    // A console application that receives a folder path, scans for Excel files, skips any workbook containing charts, and converts the remaining files to PDF using Aspose.Cells' ConversionUtility. The PDF is saved beside the source file with the same base name.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate input argument
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the folder path as the first argument.");
                return;
            }

            string folderPath = args[0];

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder does not exist: {folderPath}");
                return;
            }

            // Process each file in the folder
            foreach (string filePath in Directory.GetFiles(folderPath))
            {
                try
                {
                    // Detect if the file is an Excel workbook based on its extension
                    string extension = Path.GetExtension(filePath).ToLowerInvariant();
                    // Common Excel extensions
                    if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" &&
                        extension != ".xlsb" && extension != ".ods" && extension != ".csv")
                    {
                        continue; // Skip non‑Excel files
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Determine whether any worksheet contains charts
                    bool hasChart = false;
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        if (sheet.Charts.Count > 0)
                        {
                            hasChart = true;
                            break;
                        }
                    }

                    if (hasChart)
                    {
                        Console.WriteLine($"Skipping file (contains charts): {Path.GetFileName(filePath)}");
                        continue;
                    }

                    // Build output PDF file path
                    string pdfPath = Path.Combine(
                        Path.GetDirectoryName(filePath) ?? string.Empty,
                        Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                    // Convert Excel to PDF using the provided utility method
                    ConversionUtility.Convert(filePath, pdfPath);

                    Console.WriteLine($"Converted: {Path.GetFileName(filePath)} -> {Path.GetFileName(pdfPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }
        }
    }
}
