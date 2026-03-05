using System;
using System.Data;
using Aspose.Cells;

class ImportDataColumnDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells collection
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Prepare a DataTable with several columns (ID, Name, Score)
        DataTable dataTable = new DataTable("Students");
        dataTable.Columns.Add("ID", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Score", typeof(double));

        // Add sample rows
        dataTable.Rows.Add(1, "Alice", 85.5);
        dataTable.Rows.Add(2, "Bob", 92.0);
        dataTable.Rows.Add(3, "Charlie", 78.3);

        // Set import options to import only the "Score" column (zero‑based index 2)
        ImportTableOptions importOptions = new ImportTableOptions
        {
            ColumnIndexes = new int[] { 2 }, // import only column at index 2
            IsFieldNameShown = true          // include the column header in the sheet
        };

        // Import the selected column starting at cell A1 (row 0, column 0)
        cells.ImportData(dataTable, 0, 0, importOptions);

        // Save the workbook in XLSX format
        workbook.Save("DataColumnImport.xlsx");
    }
}