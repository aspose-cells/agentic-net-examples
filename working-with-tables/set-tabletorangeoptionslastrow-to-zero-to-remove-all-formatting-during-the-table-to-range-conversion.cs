// Title: C# – Remove All Formatting When Converting an Aspose.Cells Table to a Range (TableToRangeOptions.LastRow = 0)
// Description: Demonstrates how to set TableToRangeOptions.LastRow to 0 so that ConvertToRange strips every style from a ListObject, leaving only raw values, and saves the result as an Excel file.
// Keywords: Aspose.Cells | TableToRangeOptions | LastRow zero | convert table to range | remove formatting C# | ListObject ConvertToRange | Excel export without styles
// Common Searches: Aspose.Cells remove formatting when converting table to range | TableToRangeOptions LastRow property example | C# convert ListObject to plain range Aspose | how to strip table styles in Aspose.Cells | convert Excel table to range without styles
// Developer Intent: Use TableToRangeOptions.LastRow = 0 to discard all table formatting during a table‑to‑range conversion in Aspose.Cells.
// Use Cases: Export a styled table as plain data for downstream processing. | Reset formatting of multiple tables in a workbook before generating a clean report. | Automate batch conversion of tables to ranges while preserving only cell values.
// AI Prompts: Show a C# example that converts an Aspose.Cells ListObject to an unformatted range by setting TableToRangeOptions.LastRow to 0. | Explain the impact of TableToRangeOptions.LastRow = 0 on the worksheet after calling ConvertToRange. | Provide step‑by‑step code to create a workbook, add a table, and remove all its formatting using TableToRangeOptions.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to set TableToRangeOptions.LastRow to 0 so that ConvertToRange strips every style from a ListObject, leaving only raw values, and saves the result as an Excel file.
class TableToRangeZeroLastRowDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate the worksheet with sample data (5 columns, 10 rows)
        for (int col = 0; col < 5; col++)
        {
            cells[0, col].PutValue($"Header {col + 1}");
        }

        for (int row = 1; row < 10; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                cells[row, col].PutValue(row * col);
            }
        }

        // Add a ListObject (table) that covers the populated range
        int tableIndex = worksheet.ListObjects.Add(0, 0, 9, 4, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Create TableToRangeOptions and set LastRow to zero
        // This removes all formatting when converting the table to a range
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = 0
        };

        // Convert the table to a range using the options
        table.ConvertToRange(options);

        // Save the workbook
        workbook.Save("TableToRangeZeroLastRow.xlsx");
    }
}
