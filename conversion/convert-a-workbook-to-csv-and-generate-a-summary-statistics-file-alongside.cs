// Title: Convert an Excel workbook to individual CSV files for each worksheet and generate a summary statistics report using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export each worksheet of an .xlsx to separate CSV files | How to create a workbook summary (total rows, columns, cells) while converting Excel to CSV in .NET | Save active worksheet as CSV with Aspose.Cells and write a summary text file | C# code to generate CSV files from all sheets and produce a summary report using Aspose.Cells | Aspose.Cells .NET get total rows and columns across all worksheets
// Tags: export worksheets to CSV Aspose.Cells | calculate workbook total rows columns cells Aspose | write summary statistics file C# | set active worksheet before CSV save Aspose | clear print area for CSV export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookToCsvWithSummary
{
    // The example loads an .xlsx workbook with Aspose.Cells, creates an output folder, iterates through every worksheet, records the used row and column ranges to accumulate total rows, columns, and cells, activates each sheet, clears any print area, saves each sheet as an individual .csv file, and finally writes a text file containing the workbook's worksheet count and aggregated row, column, and cell statistics along with generation details.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path
            string inputFile = @"C:\Data\SourceWorkbook.xlsx";

            // Output folder for CSV files and summary
            string outputFolder = @"C:\Data\Output";

            try
            {
                // Ensure the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Create output directory if it does not exist
                Directory.CreateDirectory(outputFolder);

                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(inputFile);

                // Variables for summary statistics
                int totalWorksheets = workbook.Worksheets.Count;
                long totalRows = 0;
                long totalColumns = 0;
                long totalCells = 0;

                // Iterate through each worksheet and save as CSV
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range for accurate row/column counts
                    int maxRow = sheet.Cells.MaxDataRow;       // zero‑based index
                    int maxCol = sheet.Cells.MaxDataColumn;   // zero‑based index

                    // Update summary counters (add 1 because indexes are zero‑based)
                    totalRows += (maxRow + 1);
                    totalColumns += (maxCol + 1);
                    totalCells += (long)(maxRow + 1) * (maxCol + 1);

                    // Prepare CSV file name
                    string csvFileName = Path.Combine(outputFolder, $"{sheet.Name}.csv");

                    // Set the current worksheet as active
                    workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                    // Clear any print area restrictions
                    sheet.PageSetup.PrintArea = string.Empty;

                    try
                    {
                        // Save the active worksheet as CSV
                        workbook.Save(csvFileName, SaveFormat.CSV);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to save sheet '{sheet.Name}' as CSV: {ex.Message}");
                    }
                }

                // Build summary content
                string summaryContent =
                    $"Workbook Summary Statistics{Environment.NewLine}" +
                    $"--------------------------------{Environment.NewLine}" +
                    $"Source File: {Path.GetFileName(inputFile)}{Environment.NewLine}" +
                    $"Total Worksheets: {totalWorksheets}{Environment.NewLine}" +
                    $"Total Rows (across all sheets): {totalRows}{Environment.NewLine}" +
                    $"Total Columns (across all sheets): {totalColumns}{Environment.NewLine}" +
                    $"Total Cells (approximate): {totalCells}{Environment.NewLine}" +
                    $"CSV files generated in: {outputFolder}{Environment.NewLine}" +
                    $"Generation Time: {DateTime.Now}{Environment.NewLine}";

                // Write summary to a text file
                string summaryFilePath = Path.Combine(outputFolder, "WorkbookSummary.txt");
                File.WriteAllText(summaryFilePath, summaryContent);

                Console.WriteLine("Conversion to CSV completed.");
                Console.WriteLine("Summary statistics written to: " + summaryFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
