// Title: Log mismatched cell addresses when comparing two worksheets in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, iterates over the used range of the first two worksheets, compares each cell's string value, and writes any differing addresses and values to a text log. | Adjust the worksheet comparison to ignore case and trim whitespace before logging mismatches, using Aspose.Cells in C#. | Create a variant that exports mismatched cell details to a CSV file instead of a plain‑text log, preserving A1 addresses and values.
// Common Searches: how to compare two sheets in an Excel workbook with Aspose.Cells C# | write mismatched cell addresses to a log file using Aspose.Cells .NET | Aspose.Cells iterate over used range of multiple worksheets example | C# log differences between worksheets to a text file
// Tags: compare worksheets Aspose.Cells C# | log cell mismatches to text file | iterate used range Aspose.Cells | export mismatched cells to CSV Aspose.Cells | handle missing workbook file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, verifies that at least two worksheets exist, determines the maximum used rows and columns across the first two sheets, and then iterates through each cell in that combined range. It compares the string values of corresponding cells and writes any mismatched addresses (in A1 notation) together with their values to a log file, while handling missing files and other exceptions gracefully.
class WorksheetComparer
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string logPath = "mismatch_log.txt";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook containing the two worksheets to compare
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Ensure there are at least two worksheets
            if (workbook.Worksheets.Count < 2)
            {
                Console.WriteLine("Error: The workbook must contain at least two worksheets.");
                return;
            }

            // Get the first two worksheets (adjust indices or names as needed)
            Worksheet sheetA = workbook.Worksheets[0];
            Worksheet sheetB = workbook.Worksheets[1];

            // Determine the maximum used rows and columns across both sheets
            int maxRows = Math.Max(sheetA.Cells.MaxDataRow, sheetB.Cells.MaxDataRow);
            int maxCols = Math.Max(sheetA.Cells.MaxDataColumn, sheetB.Cells.MaxDataColumn);

            // Prepare the log file (overwrites if it already exists)
            try
            {
                using (StreamWriter logWriter = new StreamWriter(logPath, false))
                {
                    // Iterate through each cell within the used range
                    for (int row = 0; row <= maxRows; row++)
                    {
                        for (int col = 0; col <= maxCols; col++)
                        {
                            // Retrieve cell values from both worksheets
                            Cell cellA = sheetA.Cells[row, col];
                            Cell cellB = sheetB.Cells[row, col];

                            // Use string representation for comparison (handles different data types)
                            string valueA = cellA.StringValue;
                            string valueB = cellB.StringValue;

                            // If values differ, write the address to the log
                            if (!string.Equals(valueA, valueB, StringComparison.Ordinal))
                            {
                                // Convert zero‑based indices to A1 style address
                                string address = CellsHelper.CellIndexToName(row, col);
                                logWriter.WriteLine($"{address}: Sheet1=\"{valueA}\" | Sheet2=\"{valueB}\"");
                            }
                        }
                    }
                }

                Console.WriteLine($"Comparison complete. Mismatched cells logged to '{logPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write log file: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
