// Title: Import a DataTable into an Aspose.Cells worksheet at row 4 while inserting rows and shifting existing data down
// AI Prompts: Use Aspose.Cells ImportTableOptions to import a DataTable starting at row index 3, inserting new rows and preserving existing worksheet rows. | Show C# code that adds initial rows, then imports a DataTable with column headers at the fourth Excel row, shifting previous rows down. | Demonstrate how to set ShiftFirstRowDown and InsertRows flags to offset a DataTable import by three rows in a .NET workbook.
// Common Searches: Aspose.Cells C# import DataTable at specific row without overwriting existing rows | How to shift existing Excel rows down when importing a DataTable with Aspose.Cells | Set first row offset to 3 for DataTable import using ImportTableOptions in .NET | Insert rows instead of overwriting while importing DataTable into Aspose.Cells worksheet | Include column headers when importing DataTable into Excel at row 4 using Aspose.Cells
// Tags: ImportTableOptions InsertRows Aspose.Cells | shift rows down Aspose.Cells import | DataTable import with row offset C# | Excel row offset data import Aspose.Cells | preserve existing rows Aspose.Cells import

using System;
using System.Data;
using Aspose.Cells;

// The example creates a workbook, writes three initial rows, builds a DataTable, configures ImportTableOptions (InsertRows = true, ShiftFirstRowDown = true, IsFieldNameShown = true), and imports the table starting at row index 3 (fourth Excel row). Existing rows are shifted down, column headers are added, and the workbook is saved as OffsetImportDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells collection
        Workbook workbook = new Workbook();                     // create
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add some existing data that will be shifted down after import
        cells["A1"].PutValue("Existing Row 1");
        cells["A2"].PutValue("Existing Row 2");
        cells["A3"].PutValue("Existing Row 3");

        // Prepare a sample DataTable to import
        DataTable table = new DataTable("Sample");
        table.Columns.Add("ID", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add(1, "Alice");
        table.Rows.Add(2, "Bob");
        table.Rows.Add(3, "Charlie");

        // Configure import options:
        // - InsertRows = true : new rows will be inserted instead of overwriting existing ones
        // - ShiftFirstRowDown = true : the first row of the worksheet is shifted down to make space
        // - IsFieldNameShown = true : include column headers
        ImportTableOptions options = new ImportTableOptions
        {
            InsertRows = true,
            ShiftFirstRowDown = true,
            IsFieldNameShown = true
        };

        // Import the DataTable starting at row index 3 (fourth row in Excel)
        // Existing rows (A1:A3) will be moved down because InsertRows is true
        cells.ImportData(table, 3, 0, options);                // import

        // Save the workbook to a file
        workbook.Save("OffsetImportDemo.xlsx");                // save
    }
}
