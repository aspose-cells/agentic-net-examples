using System;
using System.Data;
using Aspose.Cells;

namespace CustomDataSourceDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook (must contain smart markers like &MyData.Name)
            Workbook workbook = new Workbook("Template.xlsx");

            // Create a WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Prepare sample data in a DataTable
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);
            dt.Rows.Add("Bob Johnson", 35);

            // Bind the DataTable directly to the designer with a name used in smart markers
            designer.SetDataSource("MyData", dt);

            // Process the smart markers and populate the workbook
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("Result.xlsx");
        }
    }
}