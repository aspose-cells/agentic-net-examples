using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (5 columns x 10 rows)
            for (int col = 0; col < 5; col++)
            {
                cells[0, col].PutValue($"Header {col + 1}");
                for (int row = 1; row < 10; row++)
                {
                    cells[row, col].PutValue(row * (col + 1));
                }
            }

            // Add a table that covers the populated range
            int tableIndex = sheet.ListObjects.Add(0, 0, 9, 4, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Retrieve a predefined (built‑in) table style
            TableStyleCollection styleCollection = workbook.Worksheets.TableStyles;
            TableStyle predefinedStyle = styleCollection.GetBuiltinTableStyle(TableStyleType.TableStyleMedium2);

            // Apply the predefined style to the table
            table.TableStyleName = predefinedStyle.Name;

            // Preserve the original cell formatting by re‑applying the style to the table's range
            // (ApplyStyleToRange respects existing explicit formatting where possible)
            table.ApplyStyleToRange();

            // Optionally show first/last column styling
            table.ShowTableStyleFirstColumn = true;
            table.ShowTableStyleLastColumn = true;

            // Save the workbook
            workbook.Save("PredefinedTableStylePreserved.xlsx");
        }
    }
}