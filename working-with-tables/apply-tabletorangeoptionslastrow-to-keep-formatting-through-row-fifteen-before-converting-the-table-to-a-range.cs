// Title: C# – Keep Table Formatting Through Row 15 Using TableToRangeOptions.LastRow in Aspose.Cells
// Description: Shows how to create a workbook, add a styled ListObject, set TableToRangeOptions.LastRow so rows 0‑14 retain the table style, convert the table to a plain range, and save the result.
// Keywords: Aspose.Cells | TableToRangeOptions | LastRow | C# | preserve table style | convert ListObject to range | Excel row formatting | row 15
// Common Searches: Aspose.Cells TableToRangeOptions LastRow C# example | keep table formatting after converting to range | retain style for first 15 rows Aspose.Cells | flatten ListObject without losing formatting | C# Aspose.Cells keep specific rows styled
// Developer Intent: The developer wants to flatten a ListObject into a regular range while preserving the original table styling for the first fifteen rows.
// Use Cases: Export a table to Excel where the header and the initial 15 rows keep their visual design before further manipulation. | Maintain conditional formatting on a subset of rows after removing the table structure for data processing. | Apply custom formatting to rows beyond the 15th row after the table has been converted to a plain range.
// AI Prompts: Write C# code using Aspose.Cells that converts a ListObject to a range and keeps the table style for rows up to a given index. | Explain the purpose of TableToRangeOptions.LastRow, its zero‑based indexing, and how to adjust it for different row counts. | Provide a step‑by‑step tutorial for converting a table to a range while preserving formatting for the first N rows in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Shows how to create a workbook, add a styled ListObject, set TableToRangeOptions.LastRow so rows 0‑14 retain the table style, convert the table to a plain range, and save the result.
class TableToRangeLastRowDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data: 5 columns, 20 data rows (plus header row)
        for (int col = 0; col < 5; col++)
        {
            cells[0, col].PutValue($"Header {col + 1}");
        }

        for (int row = 1; row <= 20; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                cells[row, col].PutValue($"R{row}C{col + 1}");
            }
        }

        // Add a table that initially spans rows 0‑20 and columns 0‑4
        int tableIndex = sheet.ListObjects.Add(0, 0, 20, 4, true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Configure conversion options: keep formatting through row 15 (zero‑based index 14)
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = 14   // rows 0‑14 will retain table formatting after conversion
        };

        // Convert the table to a regular range using the specified options
        table.ConvertToRange(options);

        // Save the resulting workbook
        workbook.Save("TableToRange_LastRow.xlsx");
    }
}
