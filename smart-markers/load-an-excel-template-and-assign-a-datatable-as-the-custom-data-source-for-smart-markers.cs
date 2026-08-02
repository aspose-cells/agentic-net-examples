using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Load the Excel template that contains smart markers
            Workbook workbook = new Workbook("Template.xlsx");

            // Create a WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // Prepare a DataTable that will be used as the custom data source
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("ProductID", typeof(int));
            dataTable.Columns.Add("ProductName", typeof(string));
            dataTable.Columns.Add("Price", typeof(decimal));

            // Add sample rows
            dataTable.Rows.Add(1, "Laptop", 999.99m);
            dataTable.Rows.Add(2, "Smartphone", 699.49m);
            dataTable.Rows.Add(3, "Tablet", 399.00m);

            // Set the DataTable as the data source for smart markers
            designer.SetDataSource(dataTable);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the resulting workbook
            workbook.Save("Result.xlsx");
        }
    }
}