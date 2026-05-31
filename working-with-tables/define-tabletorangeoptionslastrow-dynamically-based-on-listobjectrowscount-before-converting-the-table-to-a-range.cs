using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDynamicLastRowDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (5 columns, 10 rows)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Add a ListObject (table) that covers the populated range
            // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 9, 4, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Dynamically determine the last row index of the table.
            // EndRow returns the zero‑based index of the last row of the table (including header if present).
            // This value reflects the current number of rows in the table.
            int dynamicLastRow = table.EndRow;

            // Create conversion options and set LastRow to the dynamically obtained value
            TableToRangeOptions options = new TableToRangeOptions
            {
                LastRow = dynamicLastRow
            };

            // Convert the table to a range using the options
            table.ConvertToRange(options);

            // Save the workbook
            workbook.Save("TableToRange_DynamicLastRow.xlsx");
        }
    }
}