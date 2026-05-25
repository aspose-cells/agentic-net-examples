using System;
using Aspose.Cells;

namespace UnsupportedFormulaReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path (replace with actual path)
            string inputPath = "input.xlsx";

            // Output report file path
            string reportPath = "UnsupportedFormulasReport.xlsx";

            // Load the workbook to be analyzed
            Workbook workbook = new Workbook(inputPath);

            // Create a new workbook that will hold the report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];

            // Write header row for the report
            reportSheet.Cells["A1"].PutValue("Worksheet");
            reportSheet.Cells["B1"].PutValue("Cell");
            reportSheet.Cells["C1"].PutValue("Formula");

            int reportRow = 1; // zero‑based index; start after header

            // Iterate through each worksheet in the source workbook
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

                        // Skip empty cells
                        if (cell == null || string.IsNullOrEmpty(cell.Formula))
                            continue;

                        // Check if the cell's formula contains an unsupported (custom) function
                        if (cell.HasCustomFunction)
                        {
                            // Record the worksheet name, cell address, and the formula
                            reportSheet.Cells[reportRow, 0].PutValue(sheet.Name);
                            reportSheet.Cells[reportRow, 1].PutValue(cell.Name); // e.g., "B2"
                            reportSheet.Cells[reportRow, 2].PutValue(cell.Formula);
                            reportRow++;
                        }
                    }
                }
            }

            // Auto‑fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the report workbook
            reportWorkbook.Save(reportPath);
        }
    }
}