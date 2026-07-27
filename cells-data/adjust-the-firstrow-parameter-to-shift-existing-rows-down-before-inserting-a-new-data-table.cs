using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsShiftFirstRowDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add some existing data that we want to shift down
            cells["A1"].PutValue("Existing Header");
            cells["A2"].PutValue("Existing Value 1");
            cells["A3"].PutValue("Existing Value 2");

            // Prepare a sample DataTable to be imported
            DataTable table = new DataTable();
            table.Columns.Add("Column1");
            table.Columns.Add("Column2");
            // First row of the table (will become the first row after import)
            table.Rows.Add("NewHeader1", "NewHeader2");
            // Data rows
            table.Rows.Add("Data1", "Data2");
            table.Rows.Add("Data3", "Data4");

            // Configure import options to shift the first row down
            ImportTableOptions importOptions = new ImportTableOptions
            {
                // When true, the first row of the imported table is placed one row below the specified start row
                ShiftFirstRowDown = true,
                // Do not show field names separately; they are already part of the DataTable rows
                IsFieldNameShown = false,
                // Insert new rows so existing rows are not overwritten
                InsertRows = true
            };

            // Import the DataTable starting at row index 0, column index 0
            // Because ShiftFirstRowDown = true, the first row of the table will be placed at A2,
            // pushing the existing data down to A3, A4, etc.
            cells.ImportData(table, 0, 0, importOptions);

            // Save the workbook to verify the result
            workbook.Save("ShiftFirstRowDownResult.xlsx");
        }
    }
}