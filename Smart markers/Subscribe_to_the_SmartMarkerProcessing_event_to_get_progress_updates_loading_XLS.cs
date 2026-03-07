using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerProgressDemo
{
    // Callback implementation to receive progress updates during smart marker processing
    public class SmartMarkerProgressCallback : ISmartMarkerCallBack
    {
        // This method is invoked for each smart marker that is processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            Console.WriteLine($"Processing - Sheet: {sheetIndex}, Row: {rowIndex}, Column: {colIndex}, Table: {tableName}, Column: {columnName}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a workbook that will act as the template with a smart marker
            Workbook templateWorkbook = new Workbook();
            Worksheet sheet = templateWorkbook.Worksheets[0];
            sheet.Name = "Employees";
            // Place a smart marker that will be expanded for each row in the data source
            sheet.Cells["A1"].PutValue("&=Employees.Name");
            sheet.Cells["B1"].PutValue("&=Employees.Age");

            // Initialize the designer with the template workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook,
                CallBack = new SmartMarkerProgressCallback()
            };

            // Prepare a sample data source (DataTable) for the smart markers
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);

            // Bind the data source to the designer
            designer.SetDataSource(dt);

            // Process the smart markers; progress will be printed by the callback
            designer.Process();

            // Save the processed workbook
            designer.Workbook.Save("ProcessedOutput.xlsx");
        }
    }
}