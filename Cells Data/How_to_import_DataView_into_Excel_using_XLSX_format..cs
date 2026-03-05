using System;
using System.Data;
using Aspose.Cells;

class ImportDataViewDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Build a sample DataTable that will be used to create a DataView
        DataTable table = new DataTable("Products");
        table.Columns.Add("ID", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Price", typeof(decimal));

        // Populate the DataTable with sample rows
        table.Rows.Add(1, "Laptop", 999.99m);
        table.Rows.Add(2, "Smartphone", 699.99m);
        table.Rows.Add(3, "Tablet", 399.99m);

        // Create a DataView from the DataTable
        DataView dataView = new DataView(table);

        // Import the DataView into the worksheet starting at cell A1 (row 0, column 0)
        cells.ImportDataView(dataView, 0, 0);

        // Save the workbook in XLSX format
        workbook.Save("DataViewImport.xlsx", SaveFormat.Xlsx);
    }
}