// Title: C# – Export Formula Errors to CSV with Cell Address, Error Type and Timestamp using Aspose.Cells
// Description: Loads an Excel workbook, calculates all formulas, scans each used cell for error values via CellRichValue, and writes a CSV file that records the cell address, the Aspose.Cells error enum, and an ISO‑8601 timestamp for every formula error. The workbook can then be saved after processing.
// Keywords: Aspose.Cells | .NET | C# | formula error log | CSV export | cell address | error type | timestamp | calculate formulas | CellRichValue | Excel automation
// Common Searches: Aspose.Cells export formula errors to CSV | C# get Excel formula error type and cell address | How to log formula errors with timestamp using Aspose.Cells | Generate error report for Excel workbook .NET | Aspose.Cells calculate formulas and detect errors
// Developer Intent: Create a CSV report that lists every formula error in a workbook, showing the cell reference, the specific error enum, and the time the report was generated.
// Use Cases: Audit large spreadsheets for invalid formulas before publishing. | Run a scheduled job that processes uploaded workbooks and stores an error‑log CSV for monitoring. | Provide end‑users a downloadable error report after their Excel files are processed.
// AI Prompts: Write C# code with Aspose.Cells that iterates all cells, detects formula errors, and writes a CSV containing cell address, error enum, and ISO‑8601 timestamp. | Extend the program to include the worksheet name in each row of the error log. | Add a command‑line option to exclude specific error types (e.g., #DIV/0!) from the CSV report.

using System;
using System.Text;
using System.IO;
using Aspose.Cells;

namespace FormulaErrorLogger
{
    // Loads an Excel workbook, calculates all formulas, scans each used cell for error values via CellRichValue, and writes a CSV file that records the cell address, the Aspose.Cells error enum, and an ISO‑8601 timestamp for every formula error. The workbook can then be saved after processing.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be processed
            string workbookPath = "input.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Calculate all formulas so that errors are evaluated
            workbook.CalculateFormula();

            // Prepare a StringBuilder for the CSV log
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine("CellAddress,ErrorType,Timestamp");

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

                        // Retrieve the rich value to check for an error
                        CellRichValue richValue = cell.GetRichValue();

                        // If the cell contains an error, log it
                        if (richValue != null && richValue.ErrorValue != null)
                        {
                            string cellAddress = cell.Name; // e.g., "A1"
                            string errorType = richValue.ErrorValue.ToString(); // Enum name
                            string timestamp = DateTime.Now.ToString("o"); // ISO 8601 format

                            logBuilder.AppendLine($"{cellAddress},{errorType},{timestamp}");
                        }
                    }
                }
            }

            // Write the log to a CSV file
            string logPath = "FormulaErrorLog.csv";
            File.WriteAllText(logPath, logBuilder.ToString());

            // Optionally, save the workbook (uses the provided save rule)
            workbook.Save("ProcessedWorkbook.xlsx");
        }
    }
}
