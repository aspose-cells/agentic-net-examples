using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsDataViewImportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Prepare a DataTable with various data types
            DataTable table = new DataTable("SampleData");
            table.Columns.Add("ProductID", typeof(int));
            table.Columns.Add("ProductName", typeof(string));
            table.Columns.Add("Price", typeof(decimal));
            table.Columns.Add("ReleaseDate", typeof(DateTime));
            table.Columns.Add("InStock", typeof(bool));

            // Add sample rows
            table.Rows.Add(101, "Laptop", 999.99m, new DateTime(2023, 5, 1), true);
            table.Rows.Add(102, "Smartphone", 699.49m, new DateTime(2023, 6, 15), false);
            table.Rows.Add(103, "Tablet", 399.00m, new DateTime(2023, 7, 20), true);

            // Create a DataView from the DataTable
            DataView dataView = new DataView(table);

            // Import the DataView into the worksheet starting at cell A1 (row 0, column 0)
            // This method preserves the original data types of each column.
            cells.ImportDataView(dataView, 0, 0);

            // Save the workbook to an XLSX file (lifecycle: save)
            workbook.Save("DataViewImportResult.xlsx", SaveFormat.Xlsx);
        }
    }
}