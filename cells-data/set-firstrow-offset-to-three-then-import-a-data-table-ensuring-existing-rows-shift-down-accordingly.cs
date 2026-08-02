// Title: Aspose.Cells C# – Import a DataTable at Excel row 4 and shift existing rows down
// Description: Demonstrates how to set a first‑row offset of 3 (Excel row 4) when importing a DataTable into a worksheet using Aspose.Cells. The example configures ImportTableOptions with InsertRows, IsFieldNameShown, and ShiftFirstRowDown so that the new table is inserted and all pre‑existing rows move down, then saves the workbook as OffsetImportDemo.xlsx.
// Keywords: Aspose.Cells | C# | ImportData | DataTable import | firstRow offset | Excel row 4 | InsertRows true | ShiftFirstRowDown | ImportTableOptions | shift rows down | worksheet example
// Common Searches: Aspose.Cells import DataTable at specific row | how to shift existing rows when importing data with Aspose.Cells | ImportTableOptions InsertRows example C# | set firstRow offset to 3 Aspose.Cells | ShiftFirstRowDown usage Aspose.Cells
// Developer Intent: Insert a DataTable starting at row 4 while automatically moving existing worksheet rows downward.
// Use Cases: Add a table beneath static header rows in a report template without overwriting them. | Insert a new data block after introductory summary rows in an existing spreadsheet. | Create a reusable worksheet where each data import begins at a fixed offset, preserving earlier content.
// AI Prompts: Write C# code that uses Aspose.Cells to import a DataTable at Excel row 5, shifting all prior rows down and including column headers. | Explain the impact of ImportTableOptions.InsertRows and ShiftFirstRowDown on worksheet data during a DataTable import. | Provide a step‑by‑step guide for importing multiple DataTables at different row offsets in the same worksheet with Aspose.Cells.

using System;
using System.Data;
using Aspose.Cells;

// Demonstrates how to set a first‑row offset of 3 (Excel row 4) when importing a DataTable into a worksheet using Aspose.Cells. The example configures ImportTableOptions with InsertRows, IsFieldNameShown, and ShiftFirstRowDown so that the new table is inserted and all pre‑existing rows move down, then saves the workbook as OffsetImportDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells collection
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add some existing data that will be shifted down after the import
        cells["A1"].PutValue("Existing Row 1");
        cells["A2"].PutValue("Existing Row 2");
        cells["A3"].PutValue("Existing Row 3");
        cells["A4"].PutValue("Existing Row 4");
        cells["A5"].PutValue("Existing Row 5");

        // Build a DataTable to be imported
        DataTable table = new DataTable("Products");
        table.Columns.Add("ID", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Price", typeof(decimal));

        table.Rows.Add(101, "Widget", 12.5m);
        table.Rows.Add(102, "Gadget", 23.0m);
        table.Rows.Add(103, "Thingamajig", 7.75m);

        // Configure import options:
        // - InsertRows = true ensures new rows are added for the imported data,
        //   causing existing rows to shift down.
        // - IsFieldNameShown = true includes column headers.
        // - ShiftFirstRowDown = true shifts the worksheet's first row down when rows are inserted.
        ImportTableOptions options = new ImportTableOptions
        {
            InsertRows = true,
            IsFieldNameShown = true,
            ShiftFirstRowDown = true
        };

        // Set the firstRow offset to three (zero‑based index, i.e., start at Excel row 4)
        int firstRowOffset = 3;
        int firstColumn = 0;

        // Import the DataTable with the specified options
        cells.ImportData(table, firstRowOffset, firstColumn, options);

        // Save the workbook
        workbook.Save("OffsetImportDemo.xlsx");
    }
}
