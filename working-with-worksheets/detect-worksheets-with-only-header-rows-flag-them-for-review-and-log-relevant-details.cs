// Title: Identify Excel worksheets that contain only a header row using Aspose.Cells for .NET and generate a review log
// AI Prompts: Write C# code with Aspose.Cells that iterates through every worksheet in a workbook and flags sheets where the only populated row is the first (header) row. | Extend the Aspose.Cells routine to append the names of header‑only worksheets to a text log file for later review. | Enhance the detection script to include the exact header row number in the console warning for each flagged worksheet.
// Common Searches: how to find Excel sheets that only have a header row using Aspose.Cells in C# | log worksheets with no data rows beyond the header in a .NET application | detect single‑row worksheets with Aspose.Cells and export sheet names to a file
// Tags: detect header‑only worksheets Aspose.Cells | log flagged sheet names .NET | scan workbook for single populated row C# | Aspose.Cells identify empty data rows Excel

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorksheetHeaderChecker
{
    // The program loads an Excel workbook with Aspose.Cells, scans each worksheet to determine if the only non‑empty row is the first (header) row, flags such worksheets, outputs a warning with the header row number, and writes the names of those sheets to a log file for review.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be examined
            string inputFilePath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFilePath))
            {
                Console.WriteLine($"Input file not found: {inputFilePath}");
                return;
            }

            try
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(inputFilePath);

                // List to keep track of worksheets that contain only header rows
                List<string> worksheetsWithOnlyHeaders = new List<string>();

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Variables to track the first and last rows that contain any data
                    int firstDataRow = -1;
                    int lastDataRow = -1;

                    // Scan all used cells to find rows with data
                    foreach (Cell cell in cells)
                    {
                        // Use Cell.Type to determine if the cell contains data
                        if (cell.Type != CellValueType.IsNull)
                        {
                            int rowIndex = cell.Row; // zero‑based index

                            if (firstDataRow == -1 || rowIndex < firstDataRow)
                                firstDataRow = rowIndex;

                            if (rowIndex > lastDataRow)
                                lastDataRow = rowIndex;
                        }
                    }

                    // If there is at least one row with data and the first and last data rows are the same,
                    // the sheet contains data only in a single row (assumed to be the header row)
                    if (firstDataRow != -1 && firstDataRow == lastDataRow)
                    {
                        worksheetsWithOnlyHeaders.Add(sheet.Name);
                        Console.WriteLine($"[Flagged] Worksheet \"{sheet.Name}\" contains only a header row (Row {firstDataRow + 1}).");
                    }
                }

                // Optional: write flagged worksheet names to a log file
                if (worksheetsWithOnlyHeaders.Count > 0)
                {
                    string logFilePath = "HeaderReviewLog.txt";
                    File.WriteAllLines(logFilePath, worksheetsWithOnlyHeaders);
                    Console.WriteLine($"Flagged worksheets have been logged to \"{logFilePath}\".");
                }
                else
                {
                    Console.WriteLine("No worksheets with only header rows were found.");
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
