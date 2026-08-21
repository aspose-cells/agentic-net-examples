// Title: Generate a Threshold‑Based Excel Report with Aspose.Cells for .NET (C#)
// Description: C# example that loads an Excel workbook, scans every cell in the first worksheet, captures numeric values that exceed a defined limit, and writes each offending cell's address and value to a new workbook saved as a threshold report.
// Keywords: Aspose.Cells C# example | Excel threshold report | numeric cell filter | cell address extraction | audit report Aspose | value exceeds limit | C# Excel data validation
// Common Searches: Aspose.Cells list cells above a value | C# generate Excel report for values over threshold | how to extract cell addresses with high numbers using Aspose | filter numeric cells in Excel with Aspose.Cells .NET | create audit sheet for outlier values in C#
// Developer Intent: Create an Excel file that enumerates all cell addresses and their numeric values that surpass a configurable threshold.
// Use Cases: Identify budget line items that exceed spending caps. | Flag sensor readings that go beyond safety limits. | Produce an inventory audit of items whose counts are above the maximum allowed. | Generate a compliance report for regulatory thresholds in financial spreadsheets.
// AI Prompts: Write C# code with Aspose.Cells to scan a worksheet and list cells whose values are greater than a user‑provided threshold. | Explain how to modify the sample to accept the threshold as a command‑line argument and handle non‑numeric cells gracefully. | Provide a step‑by‑step tutorial for building a threshold‑based Excel audit report using Aspose.Cells, including saving the result to a new file.

using System;
using Aspose.Cells;

namespace AsposeCellsThresholdReport
{
    // C# example that loads an Excel workbook, scans every cell in the first worksheet, captures numeric values that exceed a defined limit, and writes each offending cell's address and value to a new workbook saved as a threshold report.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("input.xlsx");

            // Create a new workbook for the report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];

            // Define the numeric threshold
            double threshold = 100.0;

            // Header for the report
            reportSheet.Cells[0, 0].PutValue("Cell Address");
            reportSheet.Cells[0, 1].PutValue("Value");
            int reportRow = 1;

            // Scan all cells in the first worksheet of the source workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            foreach (Cell cell in sourceSheet.Cells)
            {
                // Check if the cell contains a numeric value exceeding the threshold
                if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue > threshold)
                {
                    // Record the cell address and its value in the report sheet
                    reportSheet.Cells[reportRow, 0].PutValue(cell.Name);
                    reportSheet.Cells[reportRow, 1].PutValue(cell.DoubleValue);
                    reportRow++;
                }
            }

            // Save the report workbook
            reportWorkbook.Save("ThresholdReport.xlsx");
        }
    }
}
