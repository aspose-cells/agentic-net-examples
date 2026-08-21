// Title: Shift existing rows down while importing a DataTable with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to use ImportTableOptions (ShiftFirstRowDown = true, InsertRows = true) to prepend a DataTable to a worksheet, moving any pre‑existing rows down instead of overwriting them. The example creates a workbook, adds sample data, imports the table starting at cell A1, and saves the result as ShiftFirstRowDemo.xlsx.
// Keywords: Aspose.Cells ShiftFirstRowDown | ImportTableOptions InsertRows C# | prepend DataTable worksheet Aspose.Cells | C# ImportData without overwriting | shift rows down Aspose.Cells .NET
// Common Searches: Aspose.Cells shift rows down when importing DataTable | ImportTableOptions ShiftFirstRowDown example C# | how to insert rows instead of overwriting with Aspose.Cells | prepend table to existing worksheet Aspose.Cells .NET | C# ImportData options for preserving existing data
// Developer Intent: Add a new table to the top of a worksheet while automatically moving existing rows down.
// Use Cases: Insert a header row above an existing dataset without losing the original rows. | Prepend a fresh DataTable to a sheet that already contains records, keeping all prior content intact. | Programmatically expand a worksheet by inserting rows during data import to maintain layout consistency.
// AI Prompts: Show C# code that imports a DataTable into an Aspose.Cells worksheet using ImportTableOptions with ShiftFirstRowDown and InsertRows enabled. | Explain the effect of ShiftFirstRowDown and InsertRows on the placement of imported data in Aspose.Cells for .NET. | Provide an example of prepending a DataTable to an existing worksheet while preserving existing rows using Aspose.Cells.

using System;
using System.Data;
using Aspose.Cells;

namespace ShiftFirstRowDemo
{
    // Demonstrates how to use ImportTableOptions (ShiftFirstRowDown = true, InsertRows = true) to prepend a DataTable to a worksheet, moving any pre‑existing rows down instead of overwriting them. The example creates a workbook, adds sample data, imports the table starting at cell A1, and saves the result as ShiftFirstRowDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add some existing data that we want to shift down
            cells["A1"].PutValue("Existing Header");
            cells["A2"].PutValue("Existing Value 1");
            cells["A3"].PutValue("Existing Value 2");

            // Build a sample DataTable to be imported
            DataTable table = new DataTable();
            table.Columns.Add("Column1");
            table.Columns.Add("Column2");
            // First row of the table (will become the first row after import)
            table.Rows.Add("NewHeader1", "NewHeader2");
            // Data rows
            table.Rows.Add("Data1", "Data2");
            table.Rows.Add("Data3", "Data4");

            // Configure import options to shift the first row down
            ImportTableOptions options = new ImportTableOptions
            {
                // Do not show field names as separate header row
                IsFieldNameShown = false,
                // Shift the first row of the imported table down by one row
                ShiftFirstRowDown = true,
                // Insert new rows instead of overwriting existing ones
                InsertRows = true
            };

            // Import the DataTable starting at row 0, column 0
            // Because ShiftFirstRowDown = true, the first row of the table will be placed at A2,
            // and the existing rows (A1:A3) will be moved down accordingly.
            cells.ImportData(table, 0, 0, options);

            // Save the workbook to verify the result
            workbook.Save("ShiftFirstRowDemo.xlsx");
        }
    }
}
