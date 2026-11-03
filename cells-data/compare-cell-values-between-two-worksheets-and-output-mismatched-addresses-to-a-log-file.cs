// Title: Compare two Excel worksheets cell‑by‑cell with Aspose.Cells for .NET and write mismatched cell addresses to a log file
// Common Searches: Aspose.Cells compare two worksheets and get list of different cells in C# | C# code to log mismatched Excel cell addresses using Aspose.Cells | How to iterate over used range of two sheets and find differences with Aspose.Cells | Write cell comparison results to a text file in .NET with Aspose.Cells | Get column letter from index while comparing Excel sheets in C#
// Tags: worksheet cell value comparison Aspose.Cells | export mismatched cell addresses to log file .NET | used range traversal Aspose.Cells | convert column index to Excel column name Aspose.Cells | compare first two sheets cell by cell Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsComparison
{
    // The example loads (or creates) an Excel workbook, ensures it contains at least two worksheets, determines the maximum used rows and columns across the first two sheets, iterates through each cell comparing values (treating nulls as equal), records the addresses of cells where the values differ, and writes those mismatched addresses to a specified text log file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file containing the two worksheets to compare
            string workbookPath = @"C:\Data\ComparisonWorkbook.xlsx";

            Workbook workbook = null;

            try
            {
                // Ensure the workbook file exists; if not, create a simple placeholder workbook
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook not found at '{workbookPath}'. Creating a sample workbook with two sheets.");
                    workbook = new Workbook();
                    workbook.Worksheets.Add("Sheet1");
                    workbook.Worksheets.Add("Sheet2");
                    // Save the placeholder for future runs
                    Directory.CreateDirectory(Path.GetDirectoryName(workbookPath));
                    workbook.Save(workbookPath);
                }
                else
                {
                    workbook = new Workbook(workbookPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load or create workbook: {ex.Message}");
                return;
            }

            // Validate that at least two worksheets are present
            if (workbook.Worksheets.Count < 2)
            {
                Console.WriteLine("The workbook must contain at least two worksheets for comparison.");
                return;
            }

            // Get the first two worksheets (adjust indices or names as needed)
            Worksheet sheet1 = workbook.Worksheets[0];
            Worksheet sheet2 = workbook.Worksheets[1];

            // Determine the maximum used rows and columns across both sheets
            int maxRows = Math.Max(sheet1.Cells.MaxDataRow, sheet2.Cells.MaxDataRow);
            int maxCols = Math.Max(sheet1.Cells.MaxDataColumn, sheet2.Cells.MaxDataColumn);

            // List to hold addresses of mismatched cells
            List<string> mismatchedAddresses = new List<string>();

            // Iterate through each cell within the used range
            for (int row = 0; row <= maxRows; row++)
            {
                for (int col = 0; col <= maxCols; col++)
                {
                    // Retrieve cell values; treat empty cells as null
                    object val1 = sheet1.Cells[row, col].Value;
                    object val2 = sheet2.Cells[row, col].Value;

                    // Compare values (handle nulls and use Equals for comparison)
                    bool areEqual = (val1 == null && val2 == null) ||
                                    (val1 != null && val2 != null && val1.Equals(val2));

                    if (!areEqual)
                    {
                        // Convert column index to letter (e.g., 0 -> A)
                        string columnLetter = CellsHelper.ColumnIndexToName(col);
                        string address = $"{columnLetter}{row + 1}";
                        mismatchedAddresses.Add(address);
                    }
                }
            }

            // Output mismatched addresses to a log file
            string logFilePath = @"C:\Data\MismatchLog.txt";

            try
            {
                // Ensure the directory for the log file exists
                Directory.CreateDirectory(Path.GetDirectoryName(logFilePath));

                using (StreamWriter writer = new StreamWriter(logFilePath, false))
                {
                    writer.WriteLine("Mismatched cell addresses between Sheet1 and Sheet2:");
                    foreach (string addr in mismatchedAddresses)
                    {
                        writer.WriteLine(addr);
                    }
                }

                Console.WriteLine($"Comparison complete. {mismatchedAddresses.Count} mismatches logged to '{logFilePath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write log file: {ex.Message}");
            }
        }
    }
}
