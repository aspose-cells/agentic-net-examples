using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the Excel template that contains smart markers
        Workbook workbook = new Workbook("Template.xlsx");

        // Create a DataTable that will serve as the custom data source
        DataTable dataTable = new DataTable("Products");
        dataTable.Columns.Add("ProductID", typeof(int));
        dataTable.Columns.Add("ProductName", typeof(string));
        dataTable.Columns.Add("Price", typeof(decimal));

        // Populate the DataTable with sample data
        dataTable.Rows.Add(1, "Laptop", 1200.50m);
        dataTable.Rows.Add(2, "Phone", 799.99m);
        dataTable.Rows.Add(3, "Tablet", 450.75m);

        // Initialize the WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Assign the DataTable as the data source for the smart markers
        designer.SetDataSource(dataTable);

        // Process the smart markers to populate the worksheet with data
        designer.Process();

        // Save the processed workbook to a new file
        workbook.Save("Result.xlsx");
    }
}