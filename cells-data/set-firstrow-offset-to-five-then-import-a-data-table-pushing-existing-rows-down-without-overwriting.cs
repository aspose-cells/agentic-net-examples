// Title: Insert a DataTable at Excel row 6 without overwriting – Aspose.Cells C# example
// Description: Demonstrates how to set a first‑row offset of 5 (Excel row 6) and import a DataTable using Aspose.Cells for .NET while inserting rows, shifting existing content down, and preserving column headers.
// Keywords: Aspose.Cells | C# | .NET | ImportData | ImportTableOptions | InsertRows | ShiftFirstRowDown | DataTable import | row offset | Excel row 6 | worksheet row insertion
// Common Searches: Aspose.Cells import DataTable at specific row | InsertRows option to push existing rows down | Set firstRow offset to 5 in cells.ImportData | How to keep existing worksheet data when importing | C# Aspose.Cells shift rows down on import
// Developer Intent: Add a DataTable starting at row 6 and automatically insert rows so existing worksheet data moves down instead of being overwritten.
// Use Cases: Add a new product list after a pre‑existing header without losing earlier rows. | Inject a dynamically generated report section into an existing sheet while preserving prior calculations. | Create a template where data blocks are inserted at fixed positions without manual row adjustments.
// AI Prompts: Show C# code using Aspose.Cells to import a DataTable at row index 5 with InsertRows and ShiftFirstRowDown enabled. | Explain the effect of ImportTableOptions properties (InsertRows, ShiftFirstRowDown, IsFieldNameShown) on row insertion and header placement. | Generate a step‑by‑step guide for importing data with a row offset while keeping existing rows intact.

using System;
using System.Data;
using Aspose.Cells;

// Demonstrates how to set a first‑row offset of 5 (Excel row 6) and import a DataTable using Aspose.Cells for .NET while inserting rows, shifting existing content down, and preserving column headers.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Existing data that should stay in the sheet
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue("Existing 1");
        cells["A3"].PutValue("Existing 2");
        cells["A4"].PutValue("Existing 3");
        cells["A5"].PutValue("Existing 4");
        cells["A6"].PutValue("Existing 5");

        // Sample DataTable to be imported
        DataTable dt = new DataTable();
        dt.Columns.Add("Product");
        dt.Columns.Add("Price", typeof(decimal));
        dt.Rows.Add("Apple", 1.2m);
        dt.Rows.Add("Banana", 0.8m);
        dt.Rows.Add("Cherry", 2.5m);

        // Import options: insert rows so existing data is pushed down
        ImportTableOptions options = new ImportTableOptions
        {
            InsertRows = true,          // add new rows for each record
            ShiftFirstRowDown = true,   // shift the first row down when inserting
            IsFieldNameShown = true     // include column headers
        };

        // Set first row offset to 5 (zero‑based, i.e., Excel row 6) and import
        int startRow = 5;
        cells.ImportData(dt, startRow, 0, options);

        // Save the workbook
        workbook.Save("OffsetImportDemo.xlsx");
    }
}
