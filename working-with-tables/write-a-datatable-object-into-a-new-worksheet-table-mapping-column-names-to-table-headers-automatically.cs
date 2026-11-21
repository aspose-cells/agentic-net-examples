// Title: Write a DataTable to a worksheet and convert it into an Excel table with headers using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# write DataTable to a worksheet and generate a ListObject with headers | Convert a DataTable into an Excel table with automatic column header mapping using Aspose.Cells | How to apply a predefined table style when creating a ListObject from a DataTable in Aspose.Cells | Save a workbook after importing a DataTable and setting TableStyleMedium9 with Aspose.Cells .NET | Create an Excel .xlsx file from a DataTable and preserve column names as table headers in Aspose.Cells
// Tags: write DataTable to worksheet cells Aspose.Cells | add ListObject with header row Aspose.Cells | apply TableStyleMedium9 to Excel table Aspose.Cells | export DataTable as .xlsx using Aspose.Cells | auto map DataTable columns to Excel table headers Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new Workbook, adds a worksheet, builds a sample DataTable, writes column names and rows to cells starting at A1, then adds a ListObject covering the range with the header flag set to true, applies the TableStyleMedium9 style, ensures the output directory exists, and saves the file as DataTableToTable.xlsx, handling import and save errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet to the workbook
            Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];

            // -------------------------------------------------
            // Prepare a DataTable (replace with your own data)
            // -------------------------------------------------
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Date", typeof(DateTime));

            dt.Rows.Add(1, "Alice", DateTime.Now);
            dt.Rows.Add(2, "Bob", DateTime.Now.AddDays(1));
            dt.Rows.Add(3, "Charlie", DateTime.Now.AddDays(2));

            // -------------------------------------------------
            // Import the DataTable into the worksheet starting at cell A1.
            // Manual import is used to avoid version‑specific API issues.
            // -------------------------------------------------
            try
            {
                const int startRow = 0;      // Zero‑based index for the first row (A1)
                const int startColumn = 0;   // Zero‑based index for the first column (A)

                // Write headers
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    sheet.Cells[startRow, startColumn + col].PutValue(dt.Columns[col].ColumnName);
                }

                // Write data rows
                for (int row = 0; row < dt.Rows.Count; row++)
                {
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        sheet.Cells[startRow + 1 + row, startColumn + col].PutValue(dt.Rows[row][col]);
                    }
                }
            }
            catch (Exception importEx)
            {
                Console.WriteLine($"Error importing DataTable: {importEx.Message}");
                return;
            }

            // -------------------------------------------------
            // Create a table (ListObject) that maps column names to headers automatically
            // -------------------------------------------------
            int firstRow = 0;                                 // Zero‑based index for the first row (A1)
            int firstColumn = 0;                              // Zero‑based index for the first column (A)
            int totalRows = dt.Rows.Count + 1;                // +1 for the header row
            int totalColumns = dt.Columns.Count;              // Number of columns

            // Add the ListObject (table) to the worksheet.
            // The last parameter 'true' indicates that the range has headers.
            ListObject table = sheet.ListObjects[sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true)];

            // Optional: set a predefined table style
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // -------------------------------------------------
            // Save the workbook to a file
            // -------------------------------------------------
            string outputPath = "DataTableToTable.xlsx";

            // Ensure the directory exists before saving
            try
            {
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
