// Title: Convert Excel to CSV and Create a Summary Statistics File with Aspose.Cells (C#)
// Description: C# example that uses Aspose.Cells ConversionUtility to export an .xlsx workbook to CSV, then reads the workbook to calculate worksheet count, total rows, columns and non‑empty cells, builds a formatted summary and writes it asynchronously to a text file with comprehensive error handling.
// Keywords: Aspose.Cells | C# Excel to CSV | ConversionUtility | workbook summary | row count | column count | non‑empty cells | async file write | Excel statistics | CSV export
// Common Searches: Aspose.Cells convert xlsx to csv c# | C# generate workbook summary after conversion | count rows and columns in Excel using Aspose.Cells | async write text file in C# Aspose.Cells | ConversionUtility Convert method example | create CSV and summary report with Aspose.Cells
// Developer Intent: Export an Excel workbook to CSV and produce a text file summarizing its structure.
// Use Cases: Automate daily data pipelines by converting incoming Excel reports to CSV while logging sheet dimensions and non‑empty cell totals for downstream validation. | Maintain audit trails for uploaded workbooks: generate a CSV copy and a concise summary that records worksheet counts, row/column totals, and content density for compliance checks. | Integrate into CI/CD workflows to verify generated Excel files meet expected size and content thresholds before publishing, using the summary as a quick health check.
// AI Prompts: Write a C# method that uses Aspose.Cells to convert an .xlsx file to CSV and then creates a summary text file with total worksheets, rows, columns, and non‑empty cells. | Enhance the ConvertAndSummarizeAsync example with detailed exception handling for missing source files, read/write permission errors, and conversion failures. | Show how to extend the summary generation to include counts of formulas, charts, and hyperlinks per worksheet using Aspose.Cells APIs. | Provide a PowerShell script that calls the C# ConvertAndSummarizeAsync routine from the command line, passing input and output paths as arguments.

using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionAndSummary
{
    // C# example that uses Aspose.Cells ConversionUtility to export an .xlsx workbook to CSV, then reads the workbook to calculate worksheet count, total rows, columns and non‑empty cells, builds a formatted summary and writes it asynchronously to a text file with comprehensive error handling.
    public class Converter
    {
        // Converts an Excel workbook to CSV and creates a summary statistics file.
        public static async Task ConvertAndSummarizeAsync(string excelPath, string csvPath, string summaryPath)
        {
            // Ensure the source Excel file exists
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"Source file not found: {excelPath}");

            try
            {
                // ---------- Conversion ----------
                // Convert Excel to CSV using Aspose.Cells utility.
                ConversionUtility.Convert(excelPath, csvPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to convert Excel to CSV.", ex);
            }

            string summary;
            try
            {
                // ---------- Summary ----------
                // Load the workbook to gather basic statistics.
                var workbook = new Workbook(excelPath);
                var sb = new StringBuilder();

                sb.AppendLine($"Workbook: {Path.GetFileName(excelPath)}");
                sb.AppendLine($"Created on: {DateTime.Now}");
                sb.AppendLine($"Number of worksheets: {workbook.Worksheets.Count}");

                long totalRows = 0;
                long totalColumns = 0;
                long totalCellsWithData = 0;

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    var maxRow = sheet.Cells.MaxDataRow;
                    var maxCol = sheet.Cells.MaxDataColumn;

                    // MaxDataRow/Column are zero‑based; add 1 for count if data exists.
                    long rows = maxRow >= 0 ? maxRow + 1 : 0;
                    long cols = maxCol >= 0 ? maxCol + 1 : 0;

                    totalRows += rows;
                    totalColumns += cols;

                    long sheetNonEmptyCells = 0;
                    // Count non‑empty cells in the sheet.
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (cell.Value != null)
                            sheetNonEmptyCells++;
                    }

                    totalCellsWithData += sheetNonEmptyCells;

                    sb.AppendLine($"Worksheet \"{sheet.Name}\": {rows} rows, {cols} columns, {sheetNonEmptyCells} non‑empty cells");
                }

                sb.AppendLine($"Total rows (across all sheets): {totalRows}");
                sb.AppendLine($"Total columns (across all sheets): {totalColumns}");
                sb.AppendLine($"Total non‑empty cells: {totalCellsWithData}");

                summary = sb.ToString();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to generate workbook summary.", ex);
            }

            try
            {
                // Write the summary text to the specified file.
                await File.WriteAllTextAsync(summaryPath, summary);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to write summary file.", ex);
            }
        }

        // Example usage
        public static async Task Main()
        {
            string sourceExcel = "input.xlsx";          // Path to the original workbook
            string outputCsv = "output.csv";            // Desired CSV file path
            string summaryFile = "summary.txt";         // Path for the summary statistics file

            try
            {
                await ConvertAndSummarizeAsync(sourceExcel, outputCsv, summaryFile);
                Console.WriteLine("Conversion and summary generation completed successfully.");
                Console.WriteLine($"CSV file: {Path.GetFullPath(outputCsv)}");
                Console.WriteLine($"Summary file: {Path.GetFullPath(summaryFile)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
