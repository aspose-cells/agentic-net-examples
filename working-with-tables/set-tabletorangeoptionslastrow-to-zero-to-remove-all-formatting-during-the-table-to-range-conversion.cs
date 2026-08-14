// Title: Strip all table formatting by setting TableToRangeOptions.LastRow = 0 in Aspose.Cells .NET
// Description: This C# example creates a workbook, fills a 5 × 10 range, adds a styled ListObject, then converts the table to a plain range using TableToRangeOptions with LastRow set to zero. The conversion removes every table style, leaving only raw cell values, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | TableToRangeOptions | LastRow zero | remove table style | convert ListObject to range | .NET spreadsheet API | clear table formatting
// Common Searches: Aspose.Cells how to remove table style on conversion | TableToRangeOptions LastRow example C# | convert table to range without formatting Aspose | what does LastRow = 0 do in Aspose.Cells | strip table formatting programmatically
// Developer Intent: Use TableToRangeOptions with LastRow set to zero to eliminate all formatting when converting a ListObject to a regular cell range.
// Use Cases: Export a workbook where tables must appear as simple cells for compatibility with other spreadsheet tools. | Reduce file size by discarding table styles before saving. | Prepare data for downstream processing that requires unstyled cell values.
// AI Prompts: Write C# code that converts an Aspose.Cells ListObject to a range with no formatting by configuring TableToRangeOptions.LastRow = 0. | Explain the impact of assigning zero to TableToRangeOptions.LastRow on the worksheet after ConvertToRange is called. | Suggest an alternative method to clear a table's style after conversion in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableToRangeDemo
{
    // This C# example creates a workbook, fills a 5 × 10 range, adds a styled ListObject, then converts the table to a plain range using TableToRangeOptions with LastRow set to zero. The conversion removes every table style, leaving only raw cell values, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data (5 columns, 10 rows)
            for (int col = 0; col < 5; col++)
            {
                // Header row
                cells[0, col].PutValue($"Column {col + 1}");
                // Data rows
                for (int row = 1; row < 10; row++)
                {
                    cells[row, col].PutValue(row * (col + 1));
                }
            }

            // Add a ListObject (table) that covers the populated range
            int tableIndex = sheet.ListObjects.Add(0, 0, 9, 4, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Optional: apply a style to see the effect of formatting removal later
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Create TableToRangeOptions and set LastRow to zero.
            // Setting LastRow to 0 ensures that no data rows are included in the conversion,
            // effectively removing all table formatting when converting to a plain range.
            TableToRangeOptions options = new TableToRangeOptions
            {
                LastRow = 0
            };

            // Convert the table to a normal range using the options
            table.ConvertToRange(options);

            // Save the workbook to verify the result
            workbook.Save("TableToRange_NoFormatting.xlsx");
        }
    }
}
