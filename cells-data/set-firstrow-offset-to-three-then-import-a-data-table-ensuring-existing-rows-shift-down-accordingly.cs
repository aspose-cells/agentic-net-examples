// Title: Import a DataTable at row 4 and shift existing rows down with Aspose.Cells for .NET
// Description: Demonstrates how to set a first‑row offset of 3, configure ImportTableOptions (InsertRows = true, IsFieldNameShown = true), and import a DataTable so that the original rows are moved down before saving the workbook.
// Keywords: Aspose.Cells ImportData row offset | C# ImportTableOptions InsertRows | shift rows down Excel .NET | ImportData starting row 4 | Aspose.Cells DataTable import example
// Common Searches: Aspose.Cells import DataTable at specific row | how to insert rows when importing data with Aspose.Cells | C# set firstRow offset for ImportData | push existing rows down Aspose.Cells .NET
// Developer Intent: Insert a DataTable beginning at the fourth Excel row while automatically moving any pre‑existing rows lower in the sheet.
// Use Cases: Add a product table beneath a summary section without overwriting the summary. | Populate a report template that has fixed header rows, inserting new data below them. | Programmatically prepend a data set to a sheet while preserving all prior content.
// AI Prompts: Write C# code that imports a DataTable at row index 5 with InsertRows enabled and formats the header row in bold using Aspose.Cells. | Explain the effect of ImportTableOptions.InsertRows and how to change the firstRow offset for different import positions. | Show how to import several DataTables sequentially with varying offsets, keeping earlier rows intact, in Aspose.Cells for .NET.

using System;
using System.Data;
using Aspose.Cells;

// Demonstrates how to set a first‑row offset of 3, configure ImportTableOptions (InsertRows = true, IsFieldNameShown = true), and import a DataTable so that the original rows are moved down before saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add some existing data that will be shifted down after import
        cells["A1"].PutValue("Existing Row 1");
        cells["A2"].PutValue("Existing Row 2");
        cells["A3"].PutValue("Existing Row 3");

        // Prepare the DataTable to be imported
        DataTable dataTable = new DataTable("Products");
        dataTable.Columns.Add("ID", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Price", typeof(decimal));
        dataTable.Rows.Add(1, "Laptop", 999.99m);
        dataTable.Rows.Add(2, "Monitor", 199.99m);
        dataTable.Rows.Add(3, "Keyboard", 49.99m);

        // Configure import options to insert rows (shifts existing rows down)
        ImportTableOptions importOptions = new ImportTableOptions
        {
            InsertRows = true,          // ensures existing rows are shifted down
            IsFieldNameShown = true    // include column headers in the import
        };

        // Import the DataTable starting at row index 3 (fourth row in Excel)
        cells.ImportData(dataTable, 3, 0, importOptions);

        // Save the workbook
        workbook.Save("OffsetImportDemo.xlsx");
    }
}
