// Title: How to implement ISmartMarkerCallBack and attach it to WorkbookDesigner for detailed smart marker processing logs in Aspose.Cells for .NET
// AI Prompts: Create a C# class that implements Aspose.Cells.ISmartMarkerCallBack and writes the sheet index, row index, column index, table name, and column name to the console for each processed smart marker. | Show the code to assign your ISmartMarkerCallBack implementation to WorkbookDesigner.CallBack, add a smart marker to a worksheet, bind a DataTable, and invoke WorkbookDesigner.Process to trigger the callback. | Explain how to extend the callback to gather processing statistics (e.g., total markers processed) and then save the workbook after processing.
// Common Searches: asp.net how to log smart marker processing using ISmartMarkerCallBack in Aspose.Cells | example of registering a custom smart marker callback with WorkbookDesigner in C# | c# Aspose.Cells smart markers callback to get sheet, row, column details | process smart markers without preserving unknown markers using WorkbookDesigner.Process | bind DataTable to smart markers and capture processing events Aspose.Cells
// Tags: ISmartMarkerCallBack custom implementation | WorkbookDesigner callback assignment | smart marker event logging C# | DataTable binding for smart markers | process smart markers without unknown preservation

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerCallbackDemo
{
    // Implement the ISmartMarkerCallBack interface to receive processing notifications
    // The example defines a SmartMarkerCallbackDemo class that implements ISmartMarkerCallBack to output sheet, row, column, table, and column information for each smart marker. The callback is assigned to WorkbookDesigner.CallBack, a smart marker referencing a DataTable is placed in cell A1, the DataTable is bound as a data source, and WorkbookDesigner.Process is called to fire the callback. Finally, the workbook is saved as SmartMarkerCallbackDemo.xlsx.
    public class SmartMarkerCallbackDemo : ISmartMarkerCallBack
    {
        // This method is called for each smart marker that is processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            Console.WriteLine($"SmartMarker processed - Sheet: {sheetIndex}, Row: {rowIndex}, Column: {colIndex}");
            Console.WriteLine($"Table: {tableName}, Column: {columnName}");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Create a new WorkbookDesigner instance
            WorkbookDesigner designer = new WorkbookDesigner();

            // Assign a new empty workbook to the designer
            designer.Workbook = new Workbook();

            // Add a smart marker to cell A1 (reference to Table1.Column1)
            Worksheet sheet = designer.Workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("&=$Table1.Column1");

            // Register the callback implementation
            designer.CallBack = new SmartMarkerCallbackDemo();

            // Prepare a DataTable as the data source
            DataTable dataTable = new DataTable("Table1");
            dataTable.Columns.Add("Column1", typeof(string));
            dataTable.Rows.Add("First");
            dataTable.Rows.Add("Second");
            dataTable.Rows.Add("Third");

            // Bind the data source to the designer
            designer.SetDataSource(dataTable);

            // Process the smart markers (false = do not preserve unrecognized markers)
            designer.Process(false);

            // Save the resulting workbook
            designer.Workbook.Save("SmartMarkerCallbackDemo.xlsx");
        }
    }
}
