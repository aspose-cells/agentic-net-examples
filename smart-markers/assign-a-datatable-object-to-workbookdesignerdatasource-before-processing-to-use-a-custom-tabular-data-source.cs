using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsDataSourceExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (template can be empty)
            Workbook workbook = new Workbook();

            // Create a WorkbookDesigner and assign the workbook to it
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // Prepare a custom DataTable as the data source
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("ProductID", typeof(int));
            dataTable.Columns.Add("ProductName", typeof(string));
            dataTable.Columns.Add("Price", typeof(decimal));

            // Add sample rows
            dataTable.Rows.Add(1, "Laptop", 1200.00m);
            dataTable.Rows.Add(2, "Smartphone", 799.99m);
            dataTable.Rows.Add(3, "Tablet", 450.50m);

            // Set the DataTable as the data source for the designer
            designer.SetDataSource(dataTable);

            // Process smart markers (if any) using the provided data source
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("CustomDataTableOutput.xlsx");
        }
    }
}