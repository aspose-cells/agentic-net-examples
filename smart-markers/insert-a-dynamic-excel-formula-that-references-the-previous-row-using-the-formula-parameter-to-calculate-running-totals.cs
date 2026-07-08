using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsRunningTotalDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add headers for the table: ID, Value, RunningTotal
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("RunningTotal");

            // Sample data rows (ID, Value)
            int[,] data = new int[,] { { 1, 10 }, { 2, 20 }, { 3, 30 }, { 4, 25 } };
            int rows = data.GetLength(0);

            // Populate the worksheet with the sample data (starting at row 2)
            for (int i = 0; i < rows; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(data[i, 0]); // ID column (A)
                sheet.Cells[i + 1, 1].PutValue(data[i, 1]); // Value column (B)
            }

            // Create a ListObject (table) that covers the data range including headers
            // Table range: A1:C{rows+1}
            int tableIndex = sheet.ListObjects.Add(0, 0, rows, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set RunningTotal formulas:
            // Row 0 (first data row) -> RunningTotal = Value
            // Subsequent rows -> RunningTotal = PreviousRowRunningTotal + CurrentValue
            for (int i = 0; i < rows; i++)
            {
                if (i == 0)
                {
                    // First row: RunningTotal = Value (B2)
                    table.PutCellFormula(i, 2, $"=B{2 + i}");
                }
                else
                {
                    // Example formula for row i (zero‑based inside table):
                    // =C{previousRow} + B{currentRow}
                    // Table rows are offset by 1 because of the header row.
                    int currentExcelRow = i + 2; // Excel row number (1‑based)
                    int previousExcelRow = currentExcelRow - 1;
                    table.PutCellFormula(i, 2,
                        $"=C{previousExcelRow}+B{currentExcelRow}");
                }
            }

            // Calculate all formulas so that the RunningTotal column contains values
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("RunningTotalDemo.xlsx");
        }
    }
}