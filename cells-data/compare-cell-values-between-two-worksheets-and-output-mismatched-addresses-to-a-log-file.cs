// Title: Log mismatched cell addresses when comparing two worksheets in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, iterates through the used rows and columns of Sheet1 and Sheet2, compares each cell's string value, and writes the addresses of cells that differ to a text file. | Adjust the comparison to trim leading/trailing whitespace and perform a case‑insensitive string check before recording mismatched cell addresses. | Enhance the log output to include both the cell address and the differing values from each sheet, formatted like "A1: Sheet1='ValueA' vs Sheet2='ValueB'".
// Common Searches: how to compare two worksheets cell by cell with Aspose.Cells and output differences to a log file | Aspose.Cells C# find cells that differ between Sheet1 and Sheet2 and write addresses to txt | C# generate list of mismatched Excel cell addresses using Aspose.Cells library | log Excel cell value discrepancies between two sheets in .NET application
// Tags: cell comparison between worksheets Aspose.Cells | log mismatched Excel cell addresses C# | iterate used range Aspose.Cells workbook | write difference report to text file .NET | max data row column detection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads input.xlsx, compares every used cell in Sheet1 and Sheet2 with Aspose.Cells, treats null/empty as empty strings, and writes the addresses of cells whose string values differ to mismatch_log.txt.
class WorksheetComparer
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string logPath = "mismatch_log.txt";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing the two worksheets to compare
            Workbook workbook = new Workbook(inputPath);

            // Get references to the worksheets (by name)
            Worksheet sheetA = workbook.Worksheets["Sheet1"];
            Worksheet sheetB = workbook.Worksheets["Sheet2"];

            if (sheetA == null || sheetB == null)
            {
                Console.WriteLine("Error: One or both worksheets (Sheet1, Sheet2) were not found.");
                return;
            }

            // Determine the maximum used rows and columns across both sheets
            int maxRows = Math.Max(sheetA.Cells.MaxDataRow, sheetB.Cells.MaxDataRow);
            int maxCols = Math.Max(sheetA.Cells.MaxDataColumn, sheetB.Cells.MaxDataColumn);

            // Prepare the log file for mismatched cell addresses
            using (StreamWriter logWriter = new StreamWriter(logPath))
            {
                // Iterate through each cell within the used range
                for (int row = 0; row <= maxRows; row++)
                {
                    for (int col = 0; col <= maxCols; col++)
                    {
                        // Retrieve the cells from both worksheets
                        Cell cellA = sheetA.Cells[row, col];
                        Cell cellB = sheetB.Cells[row, col];

                        // Compare the cell values (treat null/empty as empty string)
                        string valueA = cellA?.StringValue ?? string.Empty;
                        string valueB = cellB?.StringValue ?? string.Empty;

                        if (!valueA.Equals(valueB, StringComparison.Ordinal))
                        {
                            // If values differ, write the address (e.g., "A1") to the log
                            string address = cellA?.Name ?? CellsHelper.CellIndexToName(row, col);
                            logWriter.WriteLine(address);
                        }
                    }
                }
            }

            Console.WriteLine($"Comparison completed. Mismatched cells are logged in \"{logPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
