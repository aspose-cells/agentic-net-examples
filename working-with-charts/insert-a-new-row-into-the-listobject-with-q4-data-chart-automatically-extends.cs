// Title: Add a Q4 row to an Excel ListObject and automatically expand linked charts with Aspose.Cells for .NET
// Common Searches: how to programmatically add a new row to an Excel table and extend its chart using Aspose.Cells C# | Aspose.Cells insert row into ListObject and auto‑update chart range in .NET | C# code to add Q4 sales data to a table and refresh linked chart with Aspose.Cells | resize Excel ListObject after inserting rows so chart grows automatically Aspose.Cells | extend chart data source after adding rows to a table using Aspose.Cells for .NET
// Tags: Excel ListObject row addition Aspose.Cells | chart data source auto‑extension after table update C# | Q4 sales entry into Excel table using Aspose.Cells | programmatic table resize triggers chart growth .NET | Aspose.Cells dynamic chart linked to ListObject

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample loads an existing workbook, checks for a ListObject, inserts a new blank row after the current data range, resizes the table to include the new row, writes Q4 sales and profit values, and saves the file. Any charts linked to the ListObject automatically expand to include the newly added data.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                return;
            }

            // Get the first ListObject (table) on the worksheet
            ListObject table = sheet.ListObjects[0];

            // Determine the row index where the new data row should be inserted (after the last data row)
            int insertRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount;

            // Insert a new blank row into the worksheet
            sheet.Cells.InsertRows(insertRowIndex, 1);

            // Resize the ListObject to include the newly inserted row.
            // The last parameter indicates whether the table has headers (true for typical tables).
            table.Resize(
                table.DataRange.FirstRow,
                table.DataRange.FirstColumn,
                table.DataRange.RowCount + 1,
                table.DataRange.ColumnCount,
                true);

            // Populate the new row with Q4 data (adjust column indexes as needed)
            sheet.Cells[insertRowIndex, table.DataRange.FirstColumn + 0].PutValue("Q4");
            sheet.Cells[insertRowIndex, table.DataRange.FirstColumn + 1].PutValue(125000); // Sales
            sheet.Cells[insertRowIndex, table.DataRange.FirstColumn + 2].PutValue(30000);  // Profit

            // Save the workbook; charts linked to the ListObject will automatically extend
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
