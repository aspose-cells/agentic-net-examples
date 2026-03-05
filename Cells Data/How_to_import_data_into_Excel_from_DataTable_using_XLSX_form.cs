using System;
using System.Data;
using Aspose.Cells;

class ImportDataTableToExcel
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells collection
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Build a DataTable with sample data
        DataTable dataTable = new DataTable("Employees");
        dataTable.Columns.Add("ID", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Department", typeof(string));

        dataTable.Rows.Add(1, "John Doe", "Engineering");
        dataTable.Rows.Add(2, "Jane Smith", "Marketing");
        dataTable.Rows.Add(3, "Mike Johnson", "Sales");

        // Configure import options – include column names as the first row in Excel
        ImportTableOptions importOptions = new ImportTableOptions();
        importOptions.IsFieldNameShown = true;

        // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
        cells.ImportData(dataTable, 0, 0, importOptions);

        // Save the workbook in XLSX format
        workbook.Save("DataTableImport.xlsx", SaveFormat.Xlsx);
    }
}