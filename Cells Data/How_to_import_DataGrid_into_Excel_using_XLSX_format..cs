using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Prepare sample data in a DataTable
        DataTable dataTable = new DataTable();
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Age", typeof(int));
        dataTable.Rows.Add("John", 30);
        dataTable.Rows.Add("Jane", 25);
        dataTable.Rows.Add("Bob", 35);

        // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
        // Add column names as headers
        int startRow = 0;
        int startColumn = 0;

        // Write headers
        for (int col = 0; col < dataTable.Columns.Count; col++)
        {
            cells[startRow, startColumn + col].PutValue(dataTable.Columns[col].ColumnName);
        }

        // Write data rows
        for (int row = 0; row < dataTable.Rows.Count; row++)
        {
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                cells[startRow + 1 + row, startColumn + col].PutValue(dataTable.Rows[row][col]);
            }
        }

        // Save the workbook in XLSX format
        workbook.Save("DataGridImportDemo.xlsx");
    }
}