// Title: Preserve Header Formatting When Converting an Aspose.Cells Table to a Range (TableToRangeOptions)
// Description: This C# example creates a workbook, adds a ListObject with a styled header row, configures TableToRangeOptions (using the LastRow property), and converts the table to a regular range while keeping the header's visual style intact. The resulting file demonstrates how to retain header formatting after conversion.
// Keywords: Aspose.Cells | TableToRangeOptions | C# | .NET | preserve header style | convert ListObject to range | Excel table conversion | header formatting retention | TableToRange example
// Common Searches: Aspose.Cells keep header formatting after ConvertToRange | TableToRangeOptions LastRow usage C# | convert Excel table to range without losing style | how to preserve table header style Aspose.Cells | C# Aspose.Cells ListObject to range conversion
// Developer Intent: Convert a ListObject to a normal range while maintaining the custom style applied to the header row.
// Use Cases: Style a table header for visual emphasis, then convert the table to a range for further processing without losing the highlight. | Generate printable reports where the header must remain formatted after the table is flattened into a regular cell range. | Automate data transformations that require table-to-range conversion while preserving header aesthetics for downstream applications.
// AI Prompts: Show me a C# code snippet that uses TableToRangeOptions to convert an Aspose.Cells ListObject to a range and keep the header row formatting. | Explain why setting the LastRow property in TableToRangeOptions preserves header styles during ConvertToRange in Aspose.Cells. | Provide step‑by‑step instructions for styling a table header and then converting the table to a regular range without losing the style, using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This C# example creates a workbook, adds a ListObject with a styled header row, configures TableToRangeOptions (using the LastRow property), and converts the table to a regular range while keeping the header's visual style intact. The resulting file demonstrates how to retain header formatting after conversion.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate data with a header row
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Score");
            for (int row = 2; row <= 5; row++)
            {
                cells[row - 1, 0].PutValue(row - 1);                     // ID
                cells[row - 1, 1].PutValue($"Person {row - 1}");        // Name
                cells[row - 1, 2].PutValue(50 + row);                   // Score
            }

            // Add a table that includes the header and data (A1:C5)
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Apply a distinct style to the header row (first row of the table)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Font.IsBold = true;

            // Determine column count of the table
            int columnCount = table.EndColumn - table.StartColumn + 1;

            // Header range: one row starting at the table's first row and spanning all columns
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange(table.StartRow, table.StartColumn, 1, columnCount);
            headerRange.SetStyle(headerStyle);

            // Configure TableToRangeOptions – setting LastRow ensures the whole table is converted
            // while the header formatting applied above is retained.
            TableToRangeOptions options = new TableToRangeOptions
            {
                LastRow = table.EndRow   // zero‑based index of the last row in the table
            };

            // Convert the table to a normal range using the options
            table.ConvertToRange(options);

            // Save the workbook
            workbook.Save("TableToRangePreserveHeader.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
