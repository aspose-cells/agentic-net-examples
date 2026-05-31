using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // 1. Create a custom DataTable and populate it with sample data
        DataTable dataTable = new DataTable("Products");
        dataTable.Columns.Add("ProductID", typeof(int));
        dataTable.Columns.Add("ProductName", typeof(string));
        dataTable.Columns.Add("Price", typeof(decimal));

        dataTable.Rows.Add(1, "Laptop", 1200.00m);
        dataTable.Rows.Add(2, "Smartphone", 800.00m);
        dataTable.Rows.Add(3, "Tablet", 450.00m);

        // 2. Initialize a new workbook and a WorkbookDesigner instance
        Workbook workbook = new Workbook();
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook; // assign the workbook to the designer

        // 3. Insert smart markers into the worksheet where data will be placed
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("&=$ProductID");     // marker for ProductID column
        sheet.Cells["B1"].PutValue("&=$ProductName");   // marker for ProductName column
        sheet.Cells["C1"].PutValue("&=$Price");         // marker for Price column

        // 4. Set the custom DataTable as the data source using the SetDataSource(DataTable) method
        designer.SetDataSource(dataTable);

        // 5. Process the smart markers to populate the worksheet with the DataTable data
        designer.Process();

        // 6. Save the resulting workbook
        workbook.Save("CustomDataTableOutput.xlsx");
    }
}