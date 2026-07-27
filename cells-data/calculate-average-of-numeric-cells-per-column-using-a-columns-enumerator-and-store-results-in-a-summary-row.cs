using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsColumnAverage
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Populate sample numeric data (5 rows, 4 columns)
            // ------------------------------------------------------------
            int dataRows = 5;
            int dataColumns = 4;
            for (int row = 0; row < dataRows; row++)
            {
                for (int col = 0; col < dataColumns; col++)
                {
                    // Example values: (row + 1) * (col + 1)
                    cells[row, col].PutValue((row + 1) * (col + 1));
                }
            }

            // ------------------------------------------------------------
            // Determine the index of the summary row (one row after data)
            // ------------------------------------------------------------
            int summaryRowIndex = dataRows; // zero‑based index

            // Optional: put a label in the first cell of the summary row
            cells[summaryRowIndex, 0].PutValue("Average");

            // ------------------------------------------------------------
            // Enumerate through each column using the Columns collection
            // ------------------------------------------------------------
            foreach (Column column in cells.Columns)
            {
                int colIndex = column.Index; // column index (zero‑based)

                double sum = 0.0;
                int count = 0;

                // Enumerate rows to collect numeric values in the current column
                IEnumerator rowEnum = cells.Rows.GetEnumerator();
                while (rowEnum.MoveNext())
                {
                    Row row = (Row)rowEnum.Current;
                    // Get the cell at the current column within this row
                    Cell cell = row[colIndex];
                    if (cell != null && cell.Type == CellValueType.IsNumeric)
                    {
                        sum += cell.DoubleValue;
                        count++;
                    }
                }

                // Calculate average; avoid division by zero
                double average = count > 0 ? sum / count : 0.0;

                // Store the average in the summary row, same column
                cells[summaryRowIndex, colIndex].PutValue(average);
            }

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("ColumnAverages.xlsx");
        }
    }
}