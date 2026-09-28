// Title: Shift existing worksheet rows down before importing a DataTable using Aspose.Cells for .NET (C#)
// AI Prompts: Configure the import options to shift the worksheet's top row and enable row insertion before importing a DataTable at the beginning of a sheet. | Generate C# code that imports a DataTable into an Excel worksheet while preserving the original first row by moving it down one position. | Demonstrate setting ImportData parameters to keep existing worksheet content intact when adding a new table with Aspose.Cells.
// Common Searches: Aspose.Cells move original first row down when importing a DataTable in C# | C# code to import a DataTable without overwriting existing cells using Aspose.Cells | How to preserve existing worksheet data while adding a new table with Aspose.Cells | Example of using InsertRows with ImportData in Aspose.Cells .NET | ImportData options to keep original rows intact in Excel using Aspose.Cells
// Tags: shift first row import Aspose.Cells | InsertRows option Aspose.Cells | ImportData preserve existing rows | DataTable import without overwriting | shift worksheet rows before import

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsShiftFirstRowDemo
{
    // The example creates a workbook, writes initial data to cells A1‑A3, builds a DataTable, and sets ImportTableOptions to shift the first row down and insert new rows. ImportData then adds the DataTable starting at the top, moving the original first row to the next row, and the workbook is saved as ShiftFirstRowDownResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Existing data that will be shifted down
            cells["A1"].PutValue("Existing Header");
            cells["A2"].PutValue("Existing Value 1");
            cells["A3"].PutValue("Existing Value 2");

            // Prepare a DataTable to be imported as a new data table
            DataTable table = new DataTable();
            table.Columns.Add("Column1");
            table.Columns.Add("Column2");
            // First row of the table (will become the first row after shifting)
            table.Rows.Add("NewHeader1", "NewHeader2");
            table.Rows.Add("NewData1", "NewData2");
            table.Rows.Add("NewData3", "NewData4");

            // Configure import options to shift the first row down
            ImportTableOptions importOptions = new ImportTableOptions
            {
                // Do not show field names separately; they are part of the DataTable rows
                IsFieldNameShown = false,
                // This property causes the existing first row (A1) to move down by one row
                ShiftFirstRowDown = true,
                // Insert new rows so existing data is not overwritten
                InsertRows = true
            };

            // Import the DataTable starting at row 0, column 0
            // Because ShiftFirstRowDown is true, the original A1 cell will move to A2,
            // and the imported data will start at A2 (first row shifted down)
            cells.ImportData(table, 0, 0, importOptions);

            // Save the workbook to verify the result
            workbook.Save("ShiftFirstRowDownResult.xlsx");
        }
    }
}
