// Title: Export a Timestamped Formula Error Log (Cell Address & Error Type) with Aspose.Cells for .NET
// Description: Loads an Excel workbook, calculates all formulas while ignoring errors, scans each worksheet's used range, extracts error details from CellRichValue, and writes a CSV‑style file that records the cell address, the error enum, and an ISO‑8601 timestamp before optionally saving the updated workbook.
// Keywords: Aspose.Cells | C# formula error logging | Excel calculation errors | CellRichValue | CalculationOptions IgnoreError | export error log CSV | timestamped error report | Aspose.Cells .NET | Excel audit | error enum
// Common Searches: Aspose.Cells log formula errors | export cell error type to CSV C# | retrieve formula error timestamp Aspose | calculate workbook ignoring errors Aspose | get error enum from CellRichValue
// Developer Intent: Create a CSV (or text) report that lists every formula error in a workbook, including the cell address, the error enum, and the detection timestamp.
// Use Cases: Generate an audit file after bulk calculations to pinpoint cells that failed. | Provide end‑users with a clear error report for troubleshooting spreadsheet issues. | Automate CI/CD quality gates that abort builds when unexpected formula errors appear. | Archive error history for regulatory compliance or audit trails. | Feed error data into monitoring dashboards to track spreadsheet health over time.
// AI Prompts: Write C# code using Aspose.Cells to calculate all formulas, ignore errors, and produce a CSV log with cell address, error enum, and ISO‑8601 timestamp. | Show how to iterate through each worksheet's used range, retrieve CellRichValue, and export formula error details to a text file. | Explain configuring CalculationOptions.IgnoreError to continue processing despite errors, then saving both the updated workbook and an error report.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaErrorLog
{
    // Loads an Excel workbook, calculates all formulas while ignoring errors, scans each worksheet's used range, extracts error details from CellRichValue, and writes a CSV‑style file that records the cell address, the error enum, and an ISO‑8601 timestamp before optionally saving the updated workbook.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string workbookPath = "input.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Calculate all formulas, storing errors in cells (ignore errors = true so calculation continues)
            CalculationOptions calcOptions = new CalculationOptions { IgnoreError = true };
            workbook.CalculateFormula(calcOptions);

            // Prepare a list to hold log entries
            List<string> errorLog = new List<string>();
            errorLog.Add("CellAddress,ErrorType,Timestamp");

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Retrieve rich value which contains error information
                        CellRichValue richValue = cell.GetRichValue();

                        // If the cell contains an error, log it
                        if (richValue != null && richValue.ErrorValue != 0)
                        {
                            string address = cell.Name; // e.g., "A1"
                            string errorType = richValue.ErrorValue.ToString(); // enum name
                            string timestamp = DateTime.Now.ToString("o"); // ISO 8601 format

                            errorLog.Add($"{address},{errorType},{timestamp}");
                        }
                    }
                }
            }

            // Write the log to a text file
            string logPath = "FormulaErrorLog.txt";
            File.WriteAllLines(logPath, errorLog);

            // Optionally, save the workbook after calculation (save rule)
            workbook.Save("output.xlsx");
        }
    }
}
