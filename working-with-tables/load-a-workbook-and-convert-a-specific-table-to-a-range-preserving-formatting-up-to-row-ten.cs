// Title: C# – Convert an Excel Table to a Range (first 10 rows) while preserving formatting with Aspose.Cells
// Description: The sample loads a workbook, accesses the initial worksheet, picks the first ListObject, applies TableToRangeOptions (LastRow = 9) to change only the top ten rows of the table into a standard range, retains all cell styles, and writes the result to a new file.
// Keywords: Aspose.Cells | C# Excel table conversion | TableToRangeOptions | convert ListObject to range | preserve cell formatting | first ten rows | Excel range example | Aspose.Cells API
// Common Searches: Aspose.Cells convert ListObject to range | C# limit table conversion to specific rows | how to keep formatting when changing Excel table to range | TableToRangeOptions LastRow usage | convert first 10 rows of Excel table with Aspose
// Developer Intent: Transform a selected table into a regular range limited to the first ten rows without losing its visual formatting.
// Use Cases: Generate a lightweight version of a report that only needs the header and first ten data rows. | Feed data to a legacy system that accepts ranges but not tables, while maintaining the original look. | Create a preview sheet where only the initial rows of a table are displayed as a plain range.
// AI Prompts: Provide C# code that converts a ListObject to a range up to row 15 and keeps all styles using Aspose.Cells. | Show how to convert multiple tables with different LastRow values in one workbook. | Explain the zero‑based indexing of TableToRangeOptions.LastRow and how to verify the conversion outcome.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample loads a workbook, accesses the initial worksheet, picks the first ListObject, applies TableToRangeOptions (LastRow = 9) to change only the top ten rows of the table into a standard range, retains all cell styles, and writes the result to a new file.
class Program
{
    static void Main()
    {
        // Load the workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one table (ListObject)
        if (sheet.ListObjects.Count == 0)
        {
            Console.WriteLine("No tables found in the worksheet.");
            return;
        }

        // Retrieve the target table (here we use the first one)
        ListObject table = sheet.ListObjects[0];

        // Define conversion options: convert only up to row 10 (zero‑based index 9)
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = 9   // rows 0‑9 correspond to the first ten rows
        };

        // Convert the table to a normal range while preserving its formatting
        table.ConvertToRange(options);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
