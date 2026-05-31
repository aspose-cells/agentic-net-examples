using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data (20 rows, 5 columns)
        for (int row = 0; row < 20; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                cells[row, col].PutValue($"Data {row}-{col}");
            }
        }

        // Add a table that initially spans rows 0‑19 and columns 0‑4
        int tableIndex = worksheet.ListObjects.Add(0, 0, 19, 4, true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Set TableToRangeOptions to keep formatting through row 15 (zero‑based index 14)
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = 14   // rows 0‑14 will be converted; rows 15‑19 remain as part of the table
        };

        // Convert the table to a range using the specified options
        table.ConvertToRange(options);

        // Save the workbook
        workbook.Save("TableToRange_With_LastRow.xlsx");
    }
}