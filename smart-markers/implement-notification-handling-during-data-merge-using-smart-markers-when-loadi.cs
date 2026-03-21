using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerNotificationDemo
{
    // Custom callback class to receive notifications during smart marker processing
    public class MySmartMarkerCallback : ISmartMarkerCallBack
    {
        // This method is called for each smart marker that is processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Example: log the processing details to the console
            Console.WriteLine($"SmartMarker processed - Sheet: {sheetIndex}, Row: {rowIndex}, Column: {colIndex}, Table: {tableName}, Column: {columnName}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook("Template.xlsx");

            // Prepare a simple data source (DataTable) for demonstration
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Assign the custom callback to receive processing notifications
            designer.CallBack = new MySmartMarkerCallback();

            // Bind the data source to a name used in the smart markers (e.g., &Employees.Name)
            designer.SetDataSource("Employees", dt);

            // Process all smart markers in the workbook
            designer.Process();

            // Save the resulting workbook
            workbook.Save("Result.xlsx");
        }
    }
}