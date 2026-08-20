// Title: Detect and Log Calculation Errors in Excel Workbooks with Aspose.Cells for .NET (C#)
// Description: Loads a workbook, forces formula evaluation, scans each worksheet's used range, identifies cells where the calculated result is an error (IsErrorValue), retrieves the specific ErrorCellValueType, and writes the sheet name, cell address, and error type to the console. Optionally saves the workbook after processing.
// Keywords: Aspose.Cells error detection C# | Excel formula error logging .NET | IsErrorValue Aspose.Cells | retrieve ErrorCellValueType | scan used range Aspose.Cells | detect #DIV/0! #REF! errors | C# Excel validation Aspose
// Common Searches: how to find cells with #DIV/0! using Aspose.Cells | Aspose.Cells iterate used range and log errors | C# detect calculation errors after CalculateFormula | retrieve error type from Aspose.Cells cell | log Excel formula errors with Aspose.Cells
// Developer Intent: Find every cell that returns a calculation error after formulas are evaluated and record its address and error type.
// Use Cases: Create an error report for a financial model that lists all #REF! and #DIV/0! cells before distribution. | Automate data‑validation during import by flagging error cells for correction in code. | Generate a diagnostic log while batch‑processing workbooks to track formula issues across sheets.
// AI Prompts: Write a C# method using Aspose.Cells that returns a dictionary where each key is a worksheet name and each value is a list of (cell address, ErrorCellValueType) tuples. | Provide code that collects all error cells into a DataTable and exports sheet, address, and error type to a CSV file. | Explain how to modify the detection loop to skip a specific error type (e.g., #N/A) while still logging other errors.

using System;
using Aspose.Cells;

namespace AsposeCellsErrorDetection
{
    // Loads a workbook, forces formula evaluation, scans each worksheet's used range, identifies cells where the calculated result is an error (IsErrorValue), retrieves the specific ErrorCellValueType, and writes the sheet name, cell address, and error type to the console. Optionally saves the workbook after processing.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure formulas are calculated so that error values are materialized
            workbook.CalculateFormula();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Check if the cell's calculated result is an error
                        if (cell.IsErrorValue)
                        {
                            // Retrieve the specific error type (optional)
                            ErrorCellValueType errorType = cell.GetRichValue().ErrorValue;

                            // Log the cell address and error type
                            Console.WriteLine($"Error in sheet '{sheet.Name}' at {cell.Name}: {errorType}");
                        }
                    }
                }
            }

            // Optionally save the workbook after processing
            workbook.Save("output.xlsx");
        }
    }
}
