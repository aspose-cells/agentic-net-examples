// Title: Set TableToRangeOptions.LastRow Dynamically from ListObject Row Count in Aspose.Cells for .NET
// Description: Shows how to read a ListObject's EndRow property, assign it to TableToRangeOptions.LastRow, convert the table to a regular range, and save the workbook. The sample uses a 5‑column table with a header and a variable number of data rows.
// Keywords: Aspose.Cells | TableToRangeOptions | LastRow | ListObject | ConvertToRange | C# | .NET | dynamic row count | Excel table conversion | Aspose.Cells example
// Common Searches: Aspose.Cells set TableToRangeOptions.LastRow | Convert ListObject to range dynamically | Get last row index of a table in Aspose.Cells C# | TableToRangeOptions LastRow based on ListObject rows | Aspose.Cells dynamic table conversion
// Developer Intent: Retrieve the actual last row index of a ListObject and assign it to TableToRangeOptions.LastRow before calling ConvertToRange.
// Use Cases: Automatically adjust the conversion range when the number of data rows changes. | Generate reports that need to convert only the populated portion of a table to a range for further processing. | Combine multiple tables where each table's exact boundaries are determined at runtime.
// AI Prompts: Write C# code using Aspose.Cells that reads a ListObject's EndRow and sets TableToRangeOptions.LastRow before converting the table to a range. | Explain how to dynamically determine the last row of an Excel table in Aspose.Cells and apply it to TableToRangeOptions for conversion. | Provide a step‑by‑step example of converting a ListObject to a range when the row count is unknown at compile time.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Shows how to read a ListObject's EndRow property, assign it to TableToRangeOptions.LastRow, convert the table to a regular range, and save the workbook. The sample uses a 5‑column table with a header and a variable number of data rows.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data (header + 9 data rows, 5 columns)
        for (int col = 0; col < 5; col++)
        {
            cells[0, col].PutValue($"Header{col + 1}");
        }

        for (int row = 1; row <= 9; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                cells[row, col].PutValue($"R{row}C{col + 1}");
            }
        }

        // Add a ListObject (table) that covers the populated range
        int tableIndex = worksheet.ListObjects.Add(0, 0, 9, 4, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Dynamically determine the last row index based on the table's row count
        // Row count includes the header row; EndRow already gives the last row index.
        int lastRowIndex = table.EndRow; // equivalent to table.StartRow + (table.EndRow - table.StartRow)

        // Set the LastRow option dynamically
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = lastRowIndex
        };

        // Convert the table to a range using the dynamically set options
        table.ConvertToRange(options);

        // Save the workbook
        workbook.Save("TableToRangeDynamicLastRow.xlsx");
    }
}
