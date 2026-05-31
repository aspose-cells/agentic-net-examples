using System;
using System.Data;
using System.Text;
using Aspose.Cells;

namespace SmartMarkerMergeLogDemo
{
    // Callback implementation that records each smart marker processing event
    public class MergeLogSmartMarkerCallback : ISmartMarkerCallBack
    {
        // StringBuilder to accumulate log entries
        public StringBuilder Log { get; } = new StringBuilder();

        // This method is called by Aspose.Cells for every smart marker processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Build a detailed log entry
            Log.AppendLine($"Processed - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:\"{tableName}\", Column:\"{columnName}\"");
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook (template) and add markers
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create empty workbook
            Worksheet sheet = workbook.Worksheets[0];               // get first worksheet

            // Insert smart markers that will be filled from the data source
            sheet.Cells["A1"].PutValue("&=Employees.Name");         // smart marker for Name column
            sheet.Cells["B1"].PutValue("&=Employees.Age");          // smart marker for Age column

            // -------------------------------------------------
            // 2. Prepare data source (DataTable in this example)
            // -------------------------------------------------
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));

            employeeTable.Rows.Add("John Doe", 30);
            employeeTable.Rows.Add("Jane Smith", 28);
            employeeTable.Rows.Add("Bob Johnson", 45);

            // -------------------------------------------------
            // 3. Set up WorkbookDesigner with callback and data source
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,                                 // assign the workbook to the designer
                CallBack = new MergeLogSmartMarkerCallback()        // attach our custom callback
            };

            // Register the data source
            designer.SetDataSource(employeeTable);

            // -------------------------------------------------
            // 4. Process smart markers (populate data)
            // -------------------------------------------------
            designer.Process();                                      // execute processing

            // -------------------------------------------------
            // 5. Retrieve and display the merge log
            // -------------------------------------------------
            var callback = (MergeLogSmartMarkerCallback)designer.CallBack;
            Console.WriteLine("=== Smart Marker Processing Log ===");
            Console.WriteLine(callback.Log.ToString());

            // -------------------------------------------------
            // 6. Save the resulting workbook
            // -------------------------------------------------
            designer.Workbook.Save("MergedOutput.xlsx");
            Console.WriteLine("Workbook saved as 'MergedOutput.xlsx'.");
        }
    }
}