// Title: Generate a CSV file listing Excel formula errors with cell addresses and timestamps using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export cells containing #DIV/0! errors to CSV with timestamps | How to log formula error values and their addresses from an Excel workbook using Aspose.Cells | Create a detailed error report for Excel formulas in .NET with cell reference and ISO 8601 timestamp | Aspose.Cells iterate worksheets and write error cells to a CSV file in C# | Save Excel formula error locations to a CSV log using Aspose.Cells for .NET
// Tags: export formula error cells to CSV with Aspose.Cells | log Excel error values and cell addresses C# | timestamped error report Aspose.Cells .NET | iterate worksheets for error detection Aspose.Cells | extract cell error type Aspose.Cells API

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, scans every worksheet and cell in the used range, identifies cells whose type is an error, captures each cell's address, the error string (e.g., "#DIV/0!"), and the current ISO‑8601 timestamp, and writes these details to a CSV file named "FormulaErrorLog.csv".
class FormulaErrorLogger
{
    static void Main()
    {
        // Path to the input workbook
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Prepare a list to hold CSV log entries (including header)
        List<string> logLines = new List<string>
        {
            "CellAddress,ErrorValue,Timestamp"
        };

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range limits
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Scan all cells within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Check if the cell contains an error value
                    if (cell.Type == CellValueType.IsError)
                    {
                        string address = cell.Name; // e.g., "B2"
                        // Retrieve the error representation (e.g., "#DIV/0!")
                        string errorValue = cell.Value?.ToString() ?? string.Empty;
                        string timestamp = DateTime.Now.ToString("o"); // ISO 8601 format

                        logLines.Add($"{address},{errorValue},{timestamp}");
                    }
                }
            }
        }

        // Path to the output CSV file
        string outputPath = "FormulaErrorLog.csv";

        try
        {
            // Write the log to a CSV file
            File.WriteAllLines(outputPath, logLines);
            Console.WriteLine($"Error log successfully written to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write log file: {ex.Message}");
        }
    }
}
