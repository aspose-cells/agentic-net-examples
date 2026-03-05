using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Prepare sample data that will be displayed in a GridView
        // ------------------------------------------------------------
        DataTable dataTable = new DataTable();
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Price", typeof(double));
        dataTable.Columns.Add("Stock", typeof(int));

        dataTable.Rows.Add("Product A", 100.5, 50);
        dataTable.Rows.Add("Product B", 200.75, 30);
        dataTable.Rows.Add("Product C", 150.25, 75);

        // ------------------------------------------------------------
        // 2. Create a new workbook and get the first worksheet
        // ------------------------------------------------------------
        Workbook workbook = new Workbook();                     // create workbook
        Worksheet worksheet = workbook.Worksheets[0];           // get first worksheet
        Cells cells = worksheet.Cells;                          // shortcut to cells collection

        // ------------------------------------------------------------
        // 3. Import the DataTable into the worksheet
        //    Starting at cell B3 => row index 2, column index 1 (zero‑based)
        //    Include column headers
        // ------------------------------------------------------------
        cells.ImportDataTable(dataTable, true, 2, 1);

        // ------------------------------------------------------------
        // 4. Save the workbook in XLSX format
        // ------------------------------------------------------------
        workbook.Save("GridViewExport.xlsx", SaveFormat.Xlsx);
    }
}