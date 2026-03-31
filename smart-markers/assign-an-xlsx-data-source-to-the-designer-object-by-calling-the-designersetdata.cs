using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsDataSourceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (template can be loaded here if needed)
            Workbook workbook = new Workbook();

            // Access the first worksheet and place smart markers where data will be inserted
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("ProductID");   // Header
            sheet.Cells["B1"].PutValue("ProductName"); // Header
            sheet.Cells["A2"].PutValue("&=$ProductID");   // Smart marker for column 1
            sheet.Cells["B2"].PutValue("&=$ProductName"); // Smart marker for column 2

            // Prepare a DataTable as the runtime data source
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("ProductID", typeof(int));
            dataTable.Columns.Add("ProductName", typeof(string));

            // Add sample rows
            dataTable.Rows.Add(101, "Laptop");
            dataTable.Rows.Add(102, "Smartphone");
            dataTable.Rows.Add(103, "Tablet");

            // Initialize the WorkbookDesigner with the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Assign the DataTable to the designer at runtime
            designer.SetDataSource(dataTable);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the resulting workbook
            workbook.Save("ProductsOutput.xlsx");
        }
    }
}