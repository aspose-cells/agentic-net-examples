using System;
using System.Data;
using System.Text;
using Aspose.Cells;

namespace SmartMarkerMergeLogDemo
{
    // Callback implementation that records each processing event
    public class MergeLogSmartMarkerCallback : ISmartMarkerCallBack
    {
        private readonly StringBuilder _logBuilder = new StringBuilder();

        // This method is called by Aspose.Cells for every smart marker cell that is processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Build a detailed log entry
            _logBuilder.AppendLine(
                $"Processed - SheetIndex: {sheetIndex}, RowIndex: {rowIndex}, ColumnIndex: {colIndex}, Table: \"{tableName}\", Column: \"{columnName}\"");
        }

        // Expose the accumulated log
        public string GetLog()
        {
            return _logBuilder.ToString();
        }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a simple workbook template containing smart markers
            // -----------------------------------------------------------------
            Workbook templateWorkbook = new Workbook();
            Worksheet ws = templateWorkbook.Worksheets[0];
            // Smart markers use the syntax &=$Table.Column
            ws.Cells["A1"].PutValue("&=$Employees.Name");
            ws.Cells["B1"].PutValue("&=$Employees.Age");
            ws.Cells["A2"].PutValue("&=$Employees.Name");
            ws.Cells["B2"].PutValue("&=$Employees.Age");

            // -----------------------------------------------------------------
            // 2. Prepare data source (DataTable) matching the smart markers
            // -----------------------------------------------------------------
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));
            employeeTable.Rows.Add("Alice", 30);
            employeeTable.Rows.Add("Bob", 25);
            employeeTable.Rows.Add("Charlie", 35);

            // -----------------------------------------------------------------
            // 3. Set up WorkbookDesigner with the callback and data source
            // -----------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook
            };

            // Instantiate the custom callback
            MergeLogSmartMarkerCallback callback = new MergeLogSmartMarkerCallback();
            designer.CallBack = callback;

            // Assign the data source
            designer.SetDataSource(employeeTable);

            // -----------------------------------------------------------------
            // 4. Process the smart markers (populate data into the worksheet)
            // -----------------------------------------------------------------
            designer.Process();

            // -----------------------------------------------------------------
            // 5. Save the resulting workbook
            // -----------------------------------------------------------------
            string outputPath = "SmartMarkerMergeLogResult.xlsx";
            designer.Workbook.Save(outputPath);

            // -----------------------------------------------------------------
            // 6. Output the detailed merge log
            // -----------------------------------------------------------------
            Console.WriteLine("=== Smart Marker Merge Log ===");
            Console.WriteLine(callback.GetLog());
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}