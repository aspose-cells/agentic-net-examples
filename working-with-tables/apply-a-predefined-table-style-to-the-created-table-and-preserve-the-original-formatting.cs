using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (5 columns x 10 rows)
            for (int col = 0; col < 5; col++)
                cells[0, col].PutValue(CellsHelper.ColumnIndexToName(col)); // header row

            for (int row = 1; row < 10; row++)
                for (int col = 0; col < 5; col++)
                    cells[row, col].PutValue(row * col);

            // Add a table that covers the populated range
            int tableIndex = sheet.ListObjects.Add(0, 0, 9, 4, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Retrieve a predefined (built‑in) table style
            TableStyleCollection styleCollection = workbook.Worksheets.TableStyles;
            TableStyle builtinStyle = styleCollection.GetBuiltinTableStyle(TableStyleType.TableStyleMedium2);

            // Apply the built‑in style to the table
            table.TableStyleName = builtinStyle.Name;
            table.ShowTableStyleFirstColumn = true;
            table.ShowTableStyleLastColumn = true;
            table.ShowTableStyleRowStripes = true;
            table.ShowTableStyleColumnStripes = true;

            // Apply the style to the table's range while preserving any existing cell formatting
            table.ApplyStyleToRange();

            // Save the workbook
            workbook.Save("PredefinedTableStyle.xlsx");
        }
    }
}