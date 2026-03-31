using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerProcessingDemo
{
    // Callback implementation to receive progress updates during smart marker processing
    public class MySmartMarkerCallback : ISmartMarkerCallBack
    {
        // This method is called for each smart marker being processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            Console.WriteLine($"Processing Sheet:{sheetIndex} Row:{rowIndex} Column:{colIndex} Table:{tableName} Column:{columnName}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook templateWorkbook = new Workbook("template.xlsx");

            // Create a WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = templateWorkbook;

            // Subscribe to smart marker processing events via the callback interface
            designer.CallBack = new MySmartMarkerCallback();

            // Prepare a simple data source (replace with actual data as needed)
            DataTable dataTable = new DataTable("Employees");
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Age", typeof(int));
            dataTable.Rows.Add("John Doe", 30);
            dataTable.Rows.Add("Jane Smith", 28);

            // Bind the data source to the designer
            designer.SetDataSource(dataTable);

            // Process the smart markers; progress will be reported through the callback
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("output.xlsx");
        }
    }
}