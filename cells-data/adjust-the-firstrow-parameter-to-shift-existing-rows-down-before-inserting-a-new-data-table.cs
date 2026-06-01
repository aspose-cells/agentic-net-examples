using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsShiftFirstRowDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Existing data that will be shifted down
            cells["A1"].PutValue("Existing Header");
            cells["A2"].PutValue("Existing Value 1");
            cells["A3"].PutValue("Existing Value 2");

            // Prepare a DataTable to import
            DataTable table = new DataTable();
            table.Columns.Add("Col1");
            table.Columns.Add("Col2");
            table.Rows.Add("NewHeader1", "NewHeader2");
            table.Rows.Add("NewData1", "NewData2");
            table.Rows.Add("NewData3", "NewData4");

            // Set import options: shift the first row down so existing rows move lower
            ImportTableOptions options = new ImportTableOptions
            {
                IsFieldNameShown = true,   // import column names as header
                ShiftFirstRowDown = true   // shift existing rows down before inserting
            };

            // Import the table starting at row 0, column 0
            // Because ShiftFirstRowDown is true, the first imported row will be placed at A2,
            // and the original rows will move down accordingly.
            cells.ImportData(table, 0, 0, options);

            // Save the workbook (lifecycle save rule)
            workbook.Save("ShiftFirstRowDemo.xlsx");
        }
    }
}