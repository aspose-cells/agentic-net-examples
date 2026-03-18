using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsColumnImportDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();                     // create
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // 2. Prepare a DataTable that contains only the column we want to import
            DataTable table = new DataTable("SampleData");
            // Define the column (e.g., "Score" of type double)
            table.Columns.Add("Score", typeof(double));

            // Add sample data to the column
            table.Rows.Add(85.5);
            table.Rows.Add(92.0);
            table.Rows.Add(78.3);
            table.Rows.Add(88.8);
            table.Rows.Add(95.2);

            // 3. Set import options to import only the first column (index 0)
            ImportTableOptions importOptions = new ImportTableOptions
            {
                // Show the column header in the first row of the worksheet
                IsFieldNameShown = true,
                // Specify which column(s) from the DataTable to import
                ColumnIndexes = new int[] { 0 },
                // Insert rows if needed (not required here but safe)
                InsertRows = true
            };

            // 4. Import the column data starting at cell A1 (row 0, column 0)
            cells.ImportData(table, 0, 0, importOptions);          // import column

            // 5. Save the workbook to an XLSX file
            workbook.Save("ImportedColumnDemo.xlsx");               // save
        }
    }
}