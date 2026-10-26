// Title: How to programmatically add rows to a named ListObject, resize the table, and refresh its linked PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add rows to a named ListObject and expand its range | Aspose.Cells refresh linked PivotTable after resizing an Excel table programmatically | Resize Table1 and update PivotTable1 using Aspose.Cells .NET | Expand a named Excel table and recalculate its pivot cache with Aspose.Cells C# | Add new data rows to a ListObject and refresh the associated pivot table in Aspose.Cells
// Tags: resize ListObject Aspose.Cells C# | refresh linked PivotTable Aspose.Cells | add rows to named Excel table programmatically | pivot cache refresh after table resize .NET | expand Excel table range Aspose.Cells | update pivot data source C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Pivot;

// The example loads a workbook, accesses a named ListObject (Table1), appends new rows, resizes the table to include them, retrieves the linked PivotTable (PivotTable1), refreshes its data cache and recalculates it, then saves the updated file.
class UpdateTableAndPivot
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet that contains the named table and pivot table
            Worksheet worksheet = workbook.Worksheets["Sheet1"]; // adjust sheet name as needed
            if (worksheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found.");
                return;
            }

            // Get the named table (ListObject) by its name
            ListObject table = worksheet.ListObjects["Table1"]; // replace with your table name
            if (table == null)
            {
                Console.WriteLine("Table 'Table1' not found.");
                return;
            }

            // Determine the current range of the table
            int startRow = table.StartRow;                     // first row of the table (includes header)
            int startColumn = table.StartColumn;               // first column of the table
            int currentDataRows = table.DataRange.RowCount;    // rows that contain data (excludes header)
            int totalColumns = table.DataRange.ColumnCount;    // number of columns in the table (including header)

            // Add new rows of data below the existing table data
            int rowsToAdd = 2; // example: adding two new rows
            for (int i = 0; i < rowsToAdd; i++)
            {
                // Calculate the target row index (after the current data rows)
                int targetRow = startRow + 1 + currentDataRows + i; // +1 to skip header row

                // Populate cells in the new row (adjust column indices and values as needed)
                worksheet.Cells[targetRow, startColumn + 0].PutValue("NewValueA");
                worksheet.Cells[targetRow, startColumn + 1].PutValue("NewValueB");
                // Add more columns if the table has more than two columns
            }

            // Resize the table to include the newly added rows (hasHeaders = true)
            int newDataRows = currentDataRows + rowsToAdd;
            int newTotalRows = 1 + newDataRows; // +1 for the header row
            table.Resize(startRow, startColumn, newTotalRows, totalColumns, true);

            // Get the linked PivotTable by its name
            PivotTable pivot = worksheet.PivotTables["PivotTable1"]; // replace with your pivot name
            if (pivot == null)
            {
                Console.WriteLine("PivotTable 'PivotTable1' not found.");
                return;
            }

            // Refresh the pivot cache and recalculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
