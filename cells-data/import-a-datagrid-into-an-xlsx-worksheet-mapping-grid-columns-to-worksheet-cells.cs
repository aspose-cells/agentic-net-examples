using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsDataGridImportDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // 2. Prepare sample data in a DataTable
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Product", typeof(string));
            dataTable.Columns.Add("Quantity", typeof(int));
            dataTable.Columns.Add("Price", typeof(decimal));
            dataTable.Columns.Add("ReleaseDate", typeof(DateTime));

            dataTable.Rows.Add("Laptop", 10, 999.99m, new DateTime(2023, 5, 1));
            dataTable.Rows.Add("Smartphone", 25, 699.50m, new DateTime(2023, 6, 15));
            dataTable.Rows.Add("Tablet", 15, 450.00m, new DateTime(2023, 7, 20));

            // 3. Manually import the DataTable into the worksheet
            int rowIndex = 0;
            int colIndex = 0;

            // Write column headers
            foreach (DataColumn column in dataTable.Columns)
            {
                cells[rowIndex, colIndex].PutValue(column.ColumnName);
                colIndex++;
            }

            // Write data rows
            rowIndex++;
            foreach (DataRow row in dataTable.Rows)
            {
                colIndex = 0;
                foreach (object value in row.ItemArray)
                {
                    cells[rowIndex, colIndex].PutValue(value);
                    colIndex++;
                }
                rowIndex++;
            }

            // 4. Save the workbook as an XLSX file
            workbook.Save("DataGridImportDemo.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("DataTable imported successfully to DataGridImportDemo.xlsx");
        }
    }
}