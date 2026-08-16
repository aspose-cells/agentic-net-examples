// Title: C# CLI Tool to Batch Convert Chart‑Free Excel Files to PDF with Aspose.Cells
// Description: A command‑line application that receives a folder path, scans for Excel workbooks, skips any workbook containing charts, and converts the remaining files to PDF using Aspose.Cells.Utility.ConversionUtility, saving the PDFs alongside the originals.
// Keywords: Aspose.Cells batch conversion | C# Excel to PDF CLI | skip charts Excel conversion | ConversionUtility example | folder processing Excel files
// Common Searches: convert all Excel files in a directory to PDF with Aspose.Cells | C# command line batch Excel to PDF conversion | skip Excel workbooks with charts when exporting to PDF | Aspose.Cells CLI tool for folder conversion | how to use ConversionUtility to save PDF from Excel
// Developer Intent: Create a console program that converts every chart‑free Excel workbook in a given folder to a PDF file.
// Use Cases: Automate nightly PDF report generation from a shared folder of spreadsheets that contain only data tables. | Add a pre‑deployment step in CI/CD pipelines to verify that chart‑free Excel inputs render correctly as PDFs. | Process user‑uploaded Excel files on a server, converting only those without charts while logging skipped items.
// AI Prompts: Generate a C# method that iterates through a directory, loads each Excel file with Aspose.Cells, detects charts, and saves chart‑free workbooks as PDF using ConversionUtility. | Extend the batch converter to write detailed logs (success, skipped, errors) to a file instead of the console. | Write unit tests for the chart‑detection routine and the PDF conversion flow of the CLI utility.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace ExcelToPdfConverter
{
    // A command‑line application that receives a folder path, scans for Excel workbooks, skips any workbook containing charts, and converts the remaining files to PDF using Aspose.Cells.Utility.ConversionUtility, saving the PDFs alongside the originals.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect a single argument: the folder path containing Excel files
            if (args.Length != 1)
            {
                Console.WriteLine("Usage: ExcelToPdfConverter <folderPath>");
                return;
            }

            string folderPath = args[0];

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Error: The folder \"{folderPath}\" does not exist.");
                return;
            }

            // Define Excel file extensions to process
            string[] excelExtensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".csv", ".tsv" };

            // Enumerate files with the defined extensions
            foreach (string filePath in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly))
            {
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (Array.IndexOf(excelExtensions, ext) < 0)
                    continue; // Skip non‑Excel files

                try
                {
                    // Load the workbook (creation rule)
                    Workbook workbook = new Workbook(filePath);

                    // Determine if any worksheet contains charts
                    bool hasCharts = false;
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        if (sheet.Charts.Count > 0)
                        {
                            hasCharts = true;
                            break;
                        }
                    }

                    if (hasCharts)
                    {
                        Console.WriteLine($"Skipping \"{Path.GetFileName(filePath)}\" because it contains charts.");
                        continue;
                    }

                    // Build the output PDF file path
                    string pdfPath = Path.Combine(Path.GetDirectoryName(filePath) ?? string.Empty,
                                                  Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                    // Convert Excel to PDF using the provided ConversionUtility (save rule)
                    ConversionUtility.Convert(filePath, pdfPath);

                    Console.WriteLine($"Converted \"{Path.GetFileName(filePath)}\" to PDF successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing \"{Path.GetFileName(filePath)}\": {ex.Message}");
                }
            }
        }
    }
}
