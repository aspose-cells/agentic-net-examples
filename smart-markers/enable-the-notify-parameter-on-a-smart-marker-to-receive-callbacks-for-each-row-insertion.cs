using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerNotifyDemo
{
    // Implement the callback interface to receive notifications for each row insertion
    public class SmartMarkerCallback : ISmartMarkerCallBack
    {
        // This method is called by Aspose.Cells for each smart marker row that is processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            Console.WriteLine($"Row inserted - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker with the Notify parameter.
            // The syntax "&=$Table.Column?Notify" tells Aspose.Cells to raise the callback for each row.
            sheet.Cells["A1"].PutValue("&=$Employees.Name?Notify");
            sheet.Cells["B1"].PutValue("&=$Employees.Age?Notify");

            // ---------- Prepare data source ----------
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("Alice", 30);
            dt.Rows.Add("Bob", 25);
            dt.Rows.Add("Charlie", 28);

            // ---------- Set up WorkbookDesigner ----------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Assign the callback implementation
                CallBack = new SmartMarkerCallback()
            };

            // Register the data source
            designer.SetDataSource(dt);

            // Process the smart markers. The callback will be invoked for each inserted row.
            designer.Process(true);

            // ---------- Save the result ----------
            workbook.Save("SmartMarkerNotifyResult.xlsx");
        }
    }
}