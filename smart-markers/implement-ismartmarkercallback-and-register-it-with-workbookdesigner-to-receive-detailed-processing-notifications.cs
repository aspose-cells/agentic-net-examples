// Title: C# – Implement ISmartMarkerCallBack and Register with WorkbookDesigner for Smart Marker Events (Aspose.Cells)
// Description: Demonstrates how to create a class that implements ISmartMarkerCallBack, attach it to WorkbookDesigner, supply a DataTable source, process smart markers, and log each marker's sheet, row, column, table, and field information before saving the workbook.
// Keywords: Aspose.Cells | ISmartMarkerCallBack | WorkbookDesigner | smart markers | .NET | C# callback example | smart marker processing events | template debugging
// Common Searches: Aspose.Cells ISmartMarkerCallBack C# example | how to register a smart marker callback with WorkbookDesigner | log smart marker processing in Aspose.Cells | receive smart marker events Aspose.Cells .NET
// Developer Intent: Add a callback to capture detailed information each time a smart marker is processed in a workbook.
// Use Cases: Debug complex smart‑marker templates by outputting processing details to the console or a log file. | Gather statistics such as markers per worksheet to monitor template usage. | Validate or modify marker metadata (table/column names) on‑the‑fly before data is written.
// AI Prompts: Show how to change SmartMarkerCallback to write processing details to a text log instead of the console. | Provide code that skips smart markers belonging to a specific table using ISmartMarkerCallBack. | Explain how to aggregate callback data into a summary report of processed smart markers.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerCallbackDemo
{
    // Implementation of the ISmartMarkerCallBack interface.
    // The Process method will be invoked for each smart marker that is processed.
    // Demonstrates how to create a class that implements ISmartMarkerCallBack, attach it to WorkbookDesigner, supply a DataTable source, process smart markers, and log each marker's sheet, row, column, table, and field information before saving the workbook.
    public class SmartMarkerCallback : ISmartMarkerCallBack
    {
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Detailed notification about the smart marker being processed.
            Console.WriteLine($"[Callback] Sheet: {sheetIndex}, Row: {rowIndex}, Column: {colIndex}, Table: {tableName}, Column: {columnName}");
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a WorkbookDesigner instance.
            WorkbookDesigner designer = new WorkbookDesigner();

            // 2. Create a new workbook and assign it to the designer.
            designer.Workbook = new Workbook();

            // 3. Add a worksheet and place a smart marker in a cell.
            Worksheet sheet = designer.Workbook.Worksheets[0];
            // Smart marker syntax: &=$Table1.Column1
            sheet.Cells["A1"].PutValue("&=$Table1.Column1");

            // 4. Register the callback to receive processing notifications.
            designer.CallBack = new SmartMarkerCallback();

            // 5. Prepare a data source (DataTable) that matches the smart marker.
            DataTable dt = new DataTable("Table1");
            dt.Columns.Add("Column1", typeof(string));
            dt.Rows.Add("First");
            dt.Rows.Add("Second");
            dt.Rows.Add("Third");

            // 6. Set the data source for the designer.
            designer.SetDataSource(dt);

            // 7. Process the smart markers. The callback will be invoked for each marker.
            designer.Process();

            // 8. Save the resulting workbook.
            designer.Workbook.Save("SmartMarkerCallbackDemo.xlsx");

            Console.WriteLine("Processing completed. Workbook saved as SmartMarkerCallbackDemo.xlsx");
        }
    }
}
