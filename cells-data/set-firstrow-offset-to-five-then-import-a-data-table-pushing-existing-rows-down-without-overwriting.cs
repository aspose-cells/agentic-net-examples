// Title: Import a DataTable into an Aspose.Cells worksheet at row 5 and insert rows to shift existing content
// AI Prompts: Use ImportData with a DataTable, startRow = 5, and ImportTableOptions.InsertRows = true to add new rows without overwriting. | Show how to set ImportTableOptions.IsFieldNameShown = false when the DataTable already includes a header row. | Demonstrate saving the workbook after importing the table with a row offset and row insertion.
// Common Searches: Aspose.Cells C# import DataTable starting at row 5 without overwriting existing rows | How to insert new rows when importing a DataTable into an Excel sheet using Aspose.Cells | ImportData with row offset and InsertRows option example in .NET | Shift existing worksheet data down while adding a DataTable in Aspose.Cells | Set firstRow parameter to 5 in ImportData method Aspose.Cells C#
// Tags: ImportData row offset Aspose.Cells | InsertRows option ImportTableOptions | DataTable to Excel worksheet C# | Preserve existing worksheet rows Aspose.Cells | ImportTableOptions IsFieldNameShown false

using System;
using System.Data;
using Aspose.Cells;

// The example creates a workbook, adds some initial rows, builds a DataTable, and then imports the table beginning at row index 5 using ImportTableOptions (InsertRows = true, IsFieldNameShown = false). This inserts new rows, pushing the original content down, and saves the result to an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells collection
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // -----------------------------------------------------------------
        // Existing data that will be shifted down when we import the table
        // -----------------------------------------------------------------
        cells["A1"].PutValue("Existing Header");
        cells["A2"].PutValue("Existing Row 1");
        cells["A3"].PutValue("Existing Row 2");

        // --------------------------------------------------------------
        // Build a sample DataTable that we want to import into the sheet
        // --------------------------------------------------------------
        DataTable table = new DataTable();
        table.Columns.Add("Column1");
        table.Columns.Add("Column2");
        table.Rows.Add("Header1", "Header2");   // optional header row
        table.Rows.Add("Data1", "Data2");
        table.Rows.Add("Data3", "Data4");

        // --------------------------------------------------------------
        // Configure import options:
        //   - InsertRows = true  => new rows are inserted, pushing existing rows down
        //   - IsFieldNameShown = false (header already in the DataTable)
        // --------------------------------------------------------------
        ImportTableOptions importOptions = new ImportTableOptions
        {
            InsertRows = true,
            IsFieldNameShown = false
        };

        // --------------------------------------------------------------
        // Import the DataTable starting at row index 5 (sixth row, zero‑based)
        // --------------------------------------------------------------
        cells.ImportData(table, 5, 0, importOptions);

        // Save the workbook to a file
        workbook.Save("ImportWithOffsetAndInsertRows.xlsx");
    }
}
