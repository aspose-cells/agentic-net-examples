// Title: Create an Excel ListObject table from a System.Data.DataTable with automatic column header mapping using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a DataTable and inserts it into the first worksheet as a ListObject, automatically using the DataTable column names as the table headers. | Show how to assign a custom display name and apply a built‑in TableStyle to the ListObject after it is created with Aspose.Cells. | Provide a complete example that saves the workbook to an .xlsx file, includes error handling, and writes a console confirmation message.
// Common Searches: how to add a ListObject to a worksheet from a DataTable using Aspose.Cells C# | Aspose.Cells C# map DataTable column names to Excel table headers automatically | apply built‑in table style to a ListObject created from a DataTable in Aspose.Cells | set display name for an Excel table created with Aspose.Cells .NET | export System.Data.DataTable to .xlsx as a formatted table with Aspose.Cells
// Tags: Aspose.Cells ListObject from DataTable | C# map DataTable columns to Excel table headers | Aspose.Cells apply built-in table style | Aspose.Cells set ListObject display name | Aspose.Cells export DataTable to .xlsx

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a sample DataTable, writes its column names and rows to the first worksheet, defines the occupied range, adds a ListObject (Excel table) with the first row as headers, assigns a display name, applies a built‑in table style, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a sample DataTable (replace with your actual data)
            DataTable dt = new DataTable("SampleData");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Date", typeof(DateTime));

            dt.Rows.Add(1, "Alice", DateTime.Now);
            dt.Rows.Add(2, "Bob", DateTime.Now.AddDays(1));
            dt.Rows.Add(3, "Charlie", DateTime.Now.AddDays(2));

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Manually import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
            // Write column headers
            for (int col = 0; col < dt.Columns.Count; col++)
            {
                worksheet.Cells[0, col].PutValue(dt.Columns[col].ColumnName);
            }

            // Write data rows
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    worksheet.Cells[row + 1, col].PutValue(dt.Rows[row][col]);
                }
            }

            // Calculate the range that the table will occupy (including header row)
            int totalRows = dt.Rows.Count + 1; // +1 for the header row
            int totalCols = dt.Columns.Count;

            // Define the cell area for the table
            CellArea tableArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = totalRows - 1,
                EndColumn = totalCols - 1
            };

            // Add a ListObject (Excel table) over the defined range.
            // The last parameter 'true' indicates that the first row contains column names.
            int tableIndex = worksheet.ListObjects.Add(
                tableArea.StartRow,
                tableArea.StartColumn,
                tableArea.EndRow,
                tableArea.EndColumn,
                true);

            // Retrieve the created table object
            ListObject table = worksheet.ListObjects[tableIndex];

            // Optional: set a display name and apply a style
            table.DisplayName = "MyDataTable";
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Save the workbook to a file
            string outputPath = "DataTableToWorksheetTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
