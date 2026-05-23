using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class WorkbookToCsvAndSummary
    {
        /// <summary>
        /// Converts an Excel workbook to CSV format and creates a simple summary statistics file.
        /// </summary>
        /// <param name="excelPath">Full path of the source Excel file.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        public static async Task RunAsync(string excelPath)
        {
            // Validate input file
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
                throw new FileNotFoundException("The specified Excel file does not exist.", excelPath);

            // Determine output file names based on the source file name
            string directory = Path.GetDirectoryName(excelPath) ?? Directory.GetCurrentDirectory();
            string baseName = Path.GetFileNameWithoutExtension(excelPath);
            string csvPath = Path.Combine(directory, $"{baseName}.csv");
            string summaryPath = Path.Combine(directory, $"{baseName}_Summary.txt");

            try
            {
                // 1. Convert Excel workbook to CSV using ConversionUtility
                ConversionUtility.Convert(excelPath, csvPath);
                Console.WriteLine($"Workbook converted to CSV: {csvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
                throw;
            }

            try
            {
                // 2. Generate simple summary statistics for the CSV file
                await Task.Run(() => GenerateCsvSummary(csvPath, summaryPath));
                Console.WriteLine($"Summary statistics written to: {summaryPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during summary generation: {ex.Message}");
                throw;
            }
        }

        // Generates a basic summary (row and column count) for a CSV file.
        private static void GenerateCsvSummary(string csvPath, string summaryPath)
        {
            if (!File.Exists(csvPath))
                throw new FileNotFoundException("CSV file not found for summary generation.", csvPath);

            int rowCount = 0;
            int columnCount = 0;

            using (var reader = new StreamReader(csvPath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    rowCount++;
                    if (rowCount == 1)
                    {
                        // Assume first row defines column count
                        columnCount = line.Split(',').Length;
                    }
                }
            }

            using (var writer = new StreamWriter(summaryPath, false))
            {
                writer.WriteLine($"Rows: {rowCount}");
                writer.WriteLine($"Columns: {columnCount}");
            }
        }

        // Example entry point
        public static async Task Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the Excel file as a command‑line argument.");
                return;
            }

            try
            {
                await RunAsync(args[0]);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}