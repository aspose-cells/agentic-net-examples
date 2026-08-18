// Title: Convert Excel Workbook to CSV and Generate Row/Column Summary with Aspose.Cells for .NET (C#)
// Description: This C# example uses Aspose.Cells to load an .xlsx workbook, asynchronously converts it to CSV via ConversionUtility, then scans the CSV to count rows and columns and writes those statistics to a plain‑text summary file. Includes basic error handling and async I/O.
// Keywords: Aspose.Cells CSV conversion C# | Excel to CSV .NET | ConversionUtility Aspose.Cells | generate CSV summary statistics | async workbook conversion | row count column count CSV | C# Excel export | Aspose.Cells example
// Common Searches: convert xlsx to csv using Aspose.Cells C# | Aspose.Cells generate CSV summary file | async Excel to CSV conversion .NET | count rows and columns in exported CSV | Aspose.Cells ConversionUtility example
// Developer Intent: The developer needs to transform an Excel workbook into a CSV file and produce a simple text file that reports the total number of rows and columns in the generated CSV.
// Use Cases: Export large Excel reports to CSV for downstream analytics while capturing dimension metadata for validation. | Automate a data‑pipeline step that creates CSV outputs and logs row/column counts for audit trails. | Schedule nightly jobs that convert workbooks to CSV and generate a quick summary to monitor data consistency.
// AI Prompts: Write a C# async method that uses Aspose.Cells ConversionUtility to convert an .xlsx file to CSV and returns the CSV path. | Create C# code that reads a CSV line by line, counts rows and columns, and writes a summary text file, optimized for large files. | Enhance the error handling and logging in the async conversion and summary generation workflow using Aspose.Cells.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionAndSummary
{
    // This C# example uses Aspose.Cells to load an .xlsx workbook, asynchronously converts it to CSV via ConversionUtility, then scans the CSV to count rows and columns and writes those statistics to a plain‑text summary file. Includes basic error handling and async I/O.
    class Program
    {
        // Entry point
        static async Task Main(string[] args)
        {
            // Example file paths (replace with actual paths as needed)
            string sourceWorkbookPath = "input.xlsx";
            string csvOutputPath = "output.csv";
            string summaryOutputPath = "summary.txt";

            try
            {
                await ConvertWorkbookToCsvAndGenerateSummaryAsync(sourceWorkbookPath, csvOutputPath, summaryOutputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        /// <param name="workbookPath">Path to the source Excel workbook.</param>
        /// <param name="csvPath">Path where the CSV file will be saved.</param>
        /// <param name="summaryPath">Path where the summary text file will be saved.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        private static async Task ConvertWorkbookToCsvAndGenerateSummaryAsync(string workbookPath, string csvPath, string summaryPath)
        {
            // Verify input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.Error.WriteLine($"Input workbook not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook (lifecycle rule)
                using (Workbook workbook = new Workbook(workbookPath))
                {
                    // Convert the workbook to CSV (feature rule)
                    // ConversionUtility handles loading and saving internally.
                    ConversionUtility.Convert(workbookPath, csvPath);
                }

                // Generate a simple summary (placeholder for AI functionality)
                await GenerateSimpleSummaryAsync(csvPath, summaryPath);

                Console.WriteLine($"Conversion completed. CSV saved to: {csvPath}");
                Console.WriteLine($"Summary generated at: {summaryPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during conversion or summary generation: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a basic summary of the CSV file (row count, column count).
        /// </summary>
        private static async Task GenerateSimpleSummaryAsync(string csvPath, string summaryPath)
        {
            if (!File.Exists(csvPath))
            {
                Console.Error.WriteLine($"CSV file not found for summary generation: {csvPath}");
                return;
            }

            try
            {
                int rowCount = 0;
                int columnCount = 0;

                using (var reader = new StreamReader(csvPath))
                {
                    string line;
                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        rowCount++;
                        if (rowCount == 1)
                        {
                            // Determine column count from header line
                            columnCount = line.Split(',').Length;
                        }
                    }
                }

                using (var writer = new StreamWriter(summaryPath, false))
                {
                    await writer.WriteLineAsync("CSV Summary");
                    await writer.WriteLineAsync($"Rows: {rowCount}");
                    await writer.WriteLineAsync($"Columns: {columnCount}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error while generating summary: {ex.Message}");
            }
        }
    }
}
