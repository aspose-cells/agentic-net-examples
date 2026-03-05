using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (default format is XLSX)
        Workbook workbook = new Workbook();

        // Access the cells collection of the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        // Prepare a DataTable with sample data to import
        DataTable dataTable = new DataTable("Products");
        dataTable.Columns.Add("ProductID", typeof(int));
        dataTable.Columns.Add("ProductName", typeof(string));
        dataTable.Columns.Add("UnitPrice", typeof(double));

        dataTable.Rows.Add(1, "Apple", 0.5);
        dataTable.Rows.Add(2, "Banana", 0.3);
        dataTable.Rows.Add(3, "Cherry", 0.8);

        // Define import options (show column headers)
        ImportTableOptions importOptions = new ImportTableOptions
        {
            IsFieldNameShown = true
        };

        // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
        cells.ImportData(dataTable, 0, 0, importOptions);

        // Save the workbook in XLSX format
        workbook.Save("ImportedData.xlsx", SaveFormat.Xlsx);
    }
}