using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class RemoveEmptyRowsWithLoop
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data with some empty rows
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Data1");
            // Row 3 left empty
            cells["A4"].PutValue("Data2");
            // Row 5 left empty
            cells["A6"].PutValue("Data3");

            // Determine the current used range
            int maxRow = cells.MaxDataRow;      // last row that contains data
            int maxCol = cells.MaxDataColumn;   // last column that contains data

            // Iterate from bottom to top to avoid index shift after deletion
            for (int row = maxRow; row >= 0; row--)
            {
                bool isEmpty = true;

                // Check each cell in the row up to the last used column
                for (int col = 0; col <= maxCol; col++)
                {
                    // If a cell has a non‑null value or a non‑empty string, the row is not empty
                    if (cells[row, col].Value != null && cells[row, col].StringValue != string.Empty)
                    {
                        isEmpty = false;
                        break;
                    }
                }

                // Delete the row if it is empty
                if (isEmpty)
                {
                    cells.DeleteRow(row);
                }
            }

            // Save the workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}