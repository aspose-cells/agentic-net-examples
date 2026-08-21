// Title: Apply Built‑In Table Style and Preserve Header Formatting with Aspose.Cells (.NET)
// Description: Creates a workbook, formats the header row with bold light‑yellow style, adds a ListObject over the data range, assigns the built‑in TableStyleMedium2, calls ApplyStyleToRange to keep the custom header intact, optionally shows first/last column styling, and saves the file as an .xlsx workbook.
// Keywords: Aspose.Cells | .NET | C# | apply built‑in table style | preserve cell formatting | ListObject TableStyleName | TableStyleMedium2 | ApplyStyleToRange | custom header style | Excel table styling
// Common Searches: Aspose.Cells apply table style without losing custom header | preserve cell styles when using ListObject table style .NET | how to keep header formatting after applying TableStyleMedium2 | ApplyStyleToRange example Aspose.Cells C# | show first and last column styling in Aspose.Cells table
// Developer Intent: Apply a predefined table style to a ListObject while retaining any manually set cell styles.
// Use Cases: Generate a worksheet, style the header row, add a table, and apply TableStyleMedium2 without overwriting the header. | Enable visual emphasis on the first and last columns of a styled table. | Export the workbook so the final file shows both the built‑in table appearance and the custom header formatting.
// AI Prompts: Show C# code that uses Aspose.Cells to apply a built‑in table style to a ListObject while preserving a custom header style. | Generate an example that sets TableStyleMedium2, calls ApplyStyleToRange, and toggles ShowTableStyleFirstColumn/ShowTableStyleLastColumn. | Explain the difference between setting TableStyleName directly and using ApplyStyleToRange for keeping existing cell formatting.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableStyleDemo
{
    // Creates a workbook, formats the header row with bold light‑yellow style, adds a ListObject over the data range, assigns the built‑in TableStyleMedium2, calls ApplyStyleToRange to keep the custom header intact, optionally shows first/last column styling, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (header + 5 rows, 4 columns)
            for (int col = 0; col < 4; col++)
            {
                cells[0, col].PutValue($"Header {col + 1}");
                // Apply a distinct style to the header to demonstrate preservation
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = System.Drawing.Color.LightYellow;
                headerStyle.Pattern = BackgroundType.Solid;
                cells[0, col].SetStyle(headerStyle);
            }

            for (int row = 1; row <= 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    cells[row, col].PutValue(row * (col + 1));
                }
            }

            // Add a table (ListObject) covering the data range
            int tableIndex = sheet.ListObjects.Add(0, 0, 5, 3, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Retrieve a built‑in table style (e.g., Medium2)
            TableStyle builtinStyle = workbook.Worksheets.TableStyles.GetBuiltinTableStyle(TableStyleType.TableStyleMedium2);

            // Apply the built‑in style to the table
            table.TableStyleName = builtinStyle.Name;

            // Preserve original formatting (e.g., header style) by applying the style only to the table range
            // without overwriting explicitly set cell styles.
            table.ApplyStyleToRange();

            // Optional: show first column style if needed
            table.ShowTableStyleFirstColumn = true;
            table.ShowTableStyleLastColumn = true;

            // Save the workbook
            workbook.Save("PredefinedTableStylePreserved.xlsx");
        }
    }
}
