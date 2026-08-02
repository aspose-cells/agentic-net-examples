using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerCallbackDemo
{
    // Implement the callback interface to receive processing notifications
    public class SmartMarkerLogger : ISmartMarkerCallBack
    {
        // This method is called for each smart marker being processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            Console.WriteLine($"SmartMarker processed - Sheet: {sheetIndex}, Row: {rowIndex}, Column: {colIndex}, Table: {tableName}, Column: {columnName}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new WorkbookDesigner instance
            WorkbookDesigner designer = new WorkbookDesigner();

            // Initialize a new workbook and assign it to the designer
            designer.Workbook = new Workbook();

            // Add a smart marker to the first worksheet (e.g., &=$Employees.Name)
            Worksheet sheet = designer.Workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("&=$Employees.Name");
            sheet.Cells["B1"].PutValue("&=$Employees.Age");

            // Register the callback to receive detailed processing notifications
            designer.CallBack = new SmartMarkerLogger();

            // Prepare a simple data source (DataTable) matching the smart marker names
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);

            // Set the data source for the designer
            designer.SetDataSource(dt);

            // Process the smart markers (populate data into the worksheet)
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("SmartMarkerCallbackResult.xlsx");
        }
    }
}