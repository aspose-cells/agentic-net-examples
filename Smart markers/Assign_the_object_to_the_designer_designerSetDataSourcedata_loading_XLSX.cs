using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    public class SetDataSourceDemo
    {
        public static void Run()
        {
            // Load the existing XLSX template
            Workbook workbook = new Workbook("template.xlsx");

            // Initialize the WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Create a sample DataTable to be used as the data source
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("ProductID", typeof(int));
            dataTable.Columns.Add("ProductName", typeof(string));
            dataTable.Columns.Add("Price", typeof(decimal));

            // Add sample rows
            dataTable.Rows.Add(1, "Laptop", 999.99m);
            dataTable.Rows.Add(2, "Smartphone", 599.49m);
            dataTable.Rows.Add(3, "Tablet", 299.00m);

            // Assign the data source to the designer
            designer.SetDataSource(dataTable);

            // Process the smart markers in the template
            designer.Process();

            // Save the populated workbook
            workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetDataSourceDemo.Run();
        }
    }
}