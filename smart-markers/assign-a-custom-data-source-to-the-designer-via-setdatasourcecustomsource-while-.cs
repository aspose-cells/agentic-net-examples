using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsCustomDataSourceDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook (XLSX file) that contains smart markers
            Workbook workbook = new Workbook("Template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Create a custom data source (DataTable) and populate it with sample data
            DataTable customSource = new DataTable("Employees");
            customSource.Columns.Add("Id", typeof(int));
            customSource.Columns.Add("Name", typeof(string));
            customSource.Columns.Add("Department", typeof(string));

            customSource.Rows.Add(1, "John Doe", "Engineering");
            customSource.Rows.Add(2, "Jane Smith", "Marketing");
            customSource.Rows.Add(3, "Mike Johnson", "Sales");

            // Assign the custom data source to the designer
            // This uses the SetDataSource(DataTable) overload
            designer.SetDataSource(customSource);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the resulting workbook to a new file
            designer.Workbook.Save("Result.xlsx");
        }
    }
}