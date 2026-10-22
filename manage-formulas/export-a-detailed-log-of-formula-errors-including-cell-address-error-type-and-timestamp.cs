// Title: Generate a CSV log of Excel formula errors with cell addresses and timestamps using Aspose.Cells for .NET
// Common Searches: how to create a CSV report of formula errors in an Excel file using Aspose.Cells C# | Aspose.Cells C# export cells containing #DIV/0! errors to CSV with address and date | log Excel calculation errors with timestamp using Aspose.Cells .NET library | C# iterate worksheets and capture error cells with Aspose.Cells | write formula error details to a CSV file with Aspose.Cells in .NET
// Tags: export formula errors to CSV Aspose.Cells | detect IsError cell type Aspose.Cells | log Excel calculation errors .NET | write error details with timestamp Aspose.Cells | iterate worksheets cells Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook with Aspose.Cells, scans every worksheet and cell, identifies cells whose value type is an error, captures each cell's address, the error string (e.g., "#DIV/0!"), and the current ISO‑8601 timestamp, then writes these records to a CSV file named FormulaErrorLog.csv.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "FormulaErrorLog.csv";

        // Verify that the input workbook exists.
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file '{inputFile}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook that needs to be inspected.
            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Prepare a list to hold CSV lines. The first line is the header.
        List<string> logLines = new List<string>
        {
            "CellAddress,ErrorType,Timestamp"
        };

        // Iterate through each worksheet in the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all cells that contain data.
            foreach (Cell cell in sheet.Cells)
            {
                // Check if the cell contains a formula error.
                if (cell.Type == CellValueType.IsError)
                {
                    // Capture the error string (e.g., "#DIV/0!").
                    string errorString = cell.StringValue;

                    // Capture the cell address (e.g., "B5").
                    string address = cell.Name;

                    // Record the current timestamp in ISO 8601 format.
                    string timestamp = DateTime.Now.ToString("o");

                    // Append the log entry.
                    logLines.Add($"{address},{errorString},{timestamp}");
                }
            }
        }

        try
        {
            // Write the collected log entries to a CSV file.
            File.WriteAllLines(outputFile, logLines);
            Console.WriteLine($"Formula error log written to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error writing log file: {ex.Message}");
        }
    }
}
