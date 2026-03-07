using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerProcessingDemo
{
    // Callback implementation to receive notifications during smart marker processing
    public class MySmartMarkerCallback : ISmartMarkerCallBack
    {
        // This method is invoked for each smart marker that is processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            Console.WriteLine($"SmartMarker processed - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the Excel template that contains smart markers
            Workbook workbook = new Workbook("Template.xlsx");   // <-- load rule

            // Create a WorkbookDesigner and associate it with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Subscribe to processing notifications via the callback interface
                CallBack = new MySmartMarkerCallback()
            };

            // Prepare a simple data source (DataTable) for demonstration
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);

            // Bind the data source to the designer
            designer.SetDataSource(dt);

            // Process the smart markers; notifications will be printed by the callback
            designer.Process(true);   // <-- process rule

            // Save the populated workbook
            workbook.Save("Result.xlsx");   // <-- save rule
        }
    }
}