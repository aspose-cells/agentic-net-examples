// Title: Aspose.Cells C# – Convert ListObject to Range, keep formatting through row 15 with TableToRangeOptions.LastRow
// Description: Shows how to set TableToRangeOptions.LastRow = 14 so rows 1‑15 keep their table style when a ListObject is converted to a regular range, then saves the workbook as an XLSX file.
// Keywords: Aspose.Cells | TableToRangeOptions | LastRow | C# convert table to range | preserve table formatting | ListObject | range conversion
// Common Searches: Aspose.Cells TableToRangeOptions.LastRow example C# | convert ListObject to range while keeping first rows formatting | keep table style for specific rows Aspose.Cells | C# Aspose.Cells convert table to range with row limit
// Developer Intent: Convert a table to a range while retaining its style for the first 15 rows.
// Use Cases: Create a report where only the top 15 rows of a table are flattened, preserving visual styling. | Run calculations on a subset of data by converting those rows to a range, leaving the rest as an editable table. | Produce a printable view of the table header and initial rows without losing the applied table theme.
// AI Prompts: Provide a C# code snippet that uses TableToRangeOptions.LastRow to convert only the first N rows of a ListObject to a range in Aspose.Cells. | Explain how TableToRangeOptions.LastRow affects formatting when a table is turned into a range with Aspose.Cells for .NET. | Show an example that converts a table to a range, keeps formatting through row 15, and leaves rows 16‑20 as a table.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableToRangeDemo
{
    // Shows how to set TableToRangeOptions.LastRow = 14 so rows 1‑15 keep their table style when a ListObject is converted to a regular range, then saves the workbook as an XLSX file.
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
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Add a table that initially covers all 20 rows (0‑based indices 0‑19)
            int tableIndex = worksheet.ListObjects.Add(0, 0, 19, 4, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Set the LastRow option to 14 (row 15 in 1‑based terms) to keep formatting up to that row
            TableToRangeOptions options = new TableToRangeOptions
            {
                LastRow = 14   // rows 0‑14 will be converted; rows 15‑19 remain as part of the original table
            };

            // Convert the table to a range using the specified options
            table.ConvertToRange(options);

            // Save the workbook
            workbook.Save("TableToRange_With_LastRow.xlsx");
        }
    }
}
